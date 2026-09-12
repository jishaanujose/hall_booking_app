import calendar
import os
import re
from datetime import date, datetime, time

import streamlit as st
from supabase import Client, create_client

st.set_page_config(page_title="HallBook", page_icon="◈", layout="wide", initial_sidebar_state="collapsed")

# Inject Custom CSS to accommodate the new multi-event list styling in the calendar grid
st.markdown("""
<style>
.calendar-grid {
    display: grid;
    grid-template-columns: repeat(7, 1fr);
    gap: 8px;
    margin-top: 20px;
}
.weekday {
    text-align: center;
    font-weight: bold;
    padding: 8px;
    background-color: #f0f2f6;
    border-radius: 4px;
}
.calendar-cell {
    border: 1px solid #e6e9ef;
    border-radius: 6px;
    min-height: 120px;
    padding: 8px;
    position: relative;
    background-color: white;
}
.calendar-cell.outside {
    background-color: #fafbfc;
    opacity: 0.5;
}
.calendar-cell.today {
    border: 2px solid #ff4b4b;
}
.day-number {
    font-weight: bold;
    font-size: 14px;
    display: block;
    margin-bottom: 6px;
}
.event-item {
    background-color: #e8f0fe;
    color: #1a73e8;
    font-size: 11px;
    padding: 2px 6px;
    border-radius: 4px;
    margin-bottom: 4px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    display: block;
}
.free-label {
    color: #28a745;
    font-size: 11px;
    font-style: italic;
}
@media (max-width: 768px) {
    .calendar-grid {
        gap: 4px; /* Shrink gaps on mobile */
    }
    .calendar-cell {
        min-height: 80px; /* Make cells slightly shorter on mobile */
        padding: 4px;
    }
    .day-number {
        font-size: 11px; /* Smaller numbers */
    }
    .event-item {
        font-size: 9px; /* Smaller event badges */
        padding: 1px 3px;
    }
    .free-label {
        font-size: 9px;
    }
}
</style>
""", unsafe_allow_html=True)


def secret(name: str) -> str:
    """Read secrets locally or from Streamlit Community Cloud."""
    return st.secrets.get(name, os.getenv(name, ""))


def client() -> Client:
    if "supabase" not in st.session_state:
        url, key = secret("SUPABASE_URL"), secret("SUPABASE_KEY")
        if not url or not key:
            st.error("Missing Supabase settings. Add SUPABASE_URL and SUPABASE_KEY to your secrets.")
            st.stop()
        st.session_state.supabase = create_client(url, key)
    return st.session_state.supabase


def current_user():
    if "user" not in st.session_state:
        return None
    return st.session_state.user


def sign_out():
    client().auth.sign_out()
    st.session_state.pop("user", None)
    st.rerun()


def format_day(value: str) -> str:
    parsed = datetime.strptime(value, "%Y-%m-%d")
    return f"{parsed.strftime('%a')}, {parsed.day} {parsed.strftime('%b %Y')}"


def parse_tsrange(tsrange_str: str):
    """
    Parses a PostgreSQL tsrange string format like '[2026-10-15 10:00:00, 2026-10-15 14:00:00)'
    into Python datetime objects and formats the time chunk.
    """
    clean = tsrange_str.replace('[', '').replace(')', '').replace('"', '')
    parts = clean.split(',')

    # Strip any time zone offsets if present to keep string parsing clean
    start_str = parts[0].split('+')[0].split('-')[0:3]
    start_str = "-".join(start_str) if len(start_str) > 3 else parts[0].split('+')[0].strip()
    end_str = parts[1].split('+')[0].strip()

    try:
        start_dt = datetime.fromisoformat(start_str)
        end_dt = datetime.fromisoformat(end_str)
    except ValueError:
        # Fallback for standard SQL timestamp formats space-separated
        start_dt = datetime.strptime(start_str.split('.')[0], "%Y-%m-%d %H:%M:%S")
        end_dt = datetime.strptime(end_str.split('.')[0], "%Y-%m-%d %H:%M:%S")

    return start_dt, end_dt


def get_bookings(year: int, month: int):
    # Query ranges overlapping with the selected month
    start = f"{year}-{month:02d}-01T00:00:00"
    last_day = calendar.monthrange(year, month)[1]
    end = f"{year}-{month:02d}-{last_day}T23:59:59"

    # Use standard overlap logic via Supabase API
    response = client().table("bookings").select(
        "id,booking_time,title,notes,contact_number,booked_by,created_at").execute()

    processed_bookings = []
    for row in (response.data or []):
        start_dt, end_dt = parse_tsrange(row["booking_time"])
        # Only keep bookings that touch the target month
        if start_dt.year == year and start_dt.month == month:
            row["booking_date_key"] = start_dt.date().isoformat()
            row["start_time_str"] = start_dt.strftime("%H:%M")
            row["end_time_str"] = end_dt.strftime("%H:%M")
            processed_bookings.append(row)

    return processed_bookings


import pandas as pd
import io


def download_excel_button(booking_data, year: int, month: int):
    """Generates an Excel download button for the active month's data."""
    if not booking_data:
        st.sidebar.info("No bookings available this month to download.")
        return

    # 1. Transform raw list of dicts into a structured list for Excel
    export_rows = []
    for b in booking_data:
        export_rows.append({
            "Booking Date": b.get("booking_date_key"),
            "Start Time": b.get("start_time_str"),
            "End Time": b.get("end_time_str"),
            "Function / Event": b.get("title"),
            "Contact Number": b.get("contact_number"),
            # "Booked By (User ID)": b.get("booked_by"),
            "Created At": b.get("created_at")
        })

    # 2. Convert to a Pandas DataFrame
    df = pd.DataFrame(export_rows)

    # 3. Write Excel tracking data into a memory buffer instead of writing a local file
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=f"Bookings-{year}-{month:02d}")

    # 4. Streamlit Download Button
    month_name = calendar.month_name[month]
    st.sidebar.download_button(
        label=f"📥 Download {month_name} Report (Excel)",
        data=buffer.getvalue(),
        file_name=f"HallBookings_{year}_{month:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

def calendar_html(year: int, month: int, bookings: list[dict]) -> str:
    # Group multiple bookings by day string key
    reserved = {}
    for entry in bookings:
        key = entry["booking_date_key"]
        if key not in reserved:
            reserved[key] = []
        reserved[key].append(entry)

    today = date.today().isoformat()
    cells = []

    for week in calendar.Calendar(firstweekday=0).monthdatescalendar(year, month):
        for day_value in week:
            key = day_value.isoformat()
            outside = day_value.month != month
            day_events = reserved.get(key, [])

            classes = "calendar-cell"
            if outside:
                classes += " outside"
            if key == today:
                classes += " today"

            label = ""
            if day_events:
                # Loop through all events booked on this day
                for event in sorted(day_events, key=lambda x: x["start_time_str"]):
                    label += f'<span class="event-item" title="{event["notes"]}">🕒 {event["start_time_str"]}-{event["end_time_str"]}: {event["title"]}</span>'
            elif not outside:
                label = '<span class="free-label">Available</span>'

            cells.append(f'<div class="{classes}"><span class="day-number">{day_value.day}</span>{label}</div>')

    return '<div class="calendar-grid"><div class="weekday">Mon</div><div class="weekday">Tue</div><div class="weekday">Wed</div><div class="weekday">Thu</div><div class="weekday">Fri</div><div class="weekday">Sat</div><div class="weekday">Sun</div>' + ''.join(
        cells) + '</div>'


@st.dialog("Reserve the hall")
def booking_dialog(user):
    st.caption("Select a future date, time slot, and add the reservation details.")
    with st.form("booking-form", clear_on_submit=True):
        booking_date = st.date_input("Date", value=date.today(), min_value=date.today())

        # Time picker additions
        t_col1, t_col2 = st.columns(2)
        with t_col1:
            start_time = st.time_input("Start Time", value=time(9, 0))
        with t_col2:
            end_time = st.time_input("End Time", value=time(17, 0))

        event = st.text_input("Event name", max_chars=100, placeholder="e.g. Family celebration")
        contact_number = st.text_input("Contact number *", max_chars=30, placeholder="e.g. +91 98765 43210")
        notes = st.text_area("Notes (optional)", max_chars=500, placeholder="Any helpful details")
        submit = st.form_submit_button("Confirm reservation", type="primary", use_container_width=True)

    if submit:
        if start_time >= end_time:
            st.error("End time must be after the start time.")
        elif not event.strip():
            st.warning("Please add an event name.")
        elif not re.fullmatch(r"[0-9+() .-]+", contact_number.strip()) or not 7 <= len(
                re.sub(r"\D", "", contact_number)) <= 20:
            st.warning("Enter a valid contact number (7–20 digits).")
        else:
            try:
                # Construct timestamps strings for postgres tsrange format
                start_timestamp = f"{booking_date.isoformat()} {start_time.isoformat()}"
                end_timestamp = f"{booking_date.isoformat()} {end_time.isoformat()}"
                tsrange_payload = f"[{start_timestamp}, {end_timestamp})"

                client().table("bookings").insert({
                    "booking_time": tsrange_payload,
                    "title": event.strip(),
                    "contact_number": contact_number.strip(),
                    "notes": notes.strip(),
                    "booked_by": user.id
                }).execute()

                st.session_state.open_booking = False
                st.rerun()
            except Exception as error:
                error_str = str(error).lower()
                # Catch the exclusion constraint error thrown by Postgres
                if "exclusion constraint" in error_str or "23p01" in error_str or "exclude" in error_str:
                    st.error("🚨 Time overlapping with another booking. Please choose a different time slot.")
                elif "23505" in error_str or "duplicate key" in error_str:
                    st.error("🚨 That slot has already been claimed.")
                else:
                    st.error(f"Something went wrong: {str(error)}")

    if st.button("Cancel", use_container_width=True):
        st.session_state.open_booking = False
        st.rerun()


def login_view():
    st.markdown('<div class="brand-lockup"><div class="brand-mark">◈</div><span>HallBook</span></div>',
                unsafe_allow_html=True)
    left, center, right = st.columns([1, 1.15, 1])
    with center:
        st.markdown(
            "<div class='login-heading'><p class='kicker'>BOOKING DETAILS</p><h1>Reserve your space<br>with confidence.</h1><p>See availability, make a reservation, and keep everyone in sync.</p></div>",
            unsafe_allow_html=True)
        tab_login, tab_create = st.tabs(["Sign in", "Create account"])
        with tab_login:
            with st.form("login"):
                email = st.text_input("Email address", placeholder="you@example.com")
                password = st.text_input("Password", type="password")
                submit = st.form_submit_button("Sign in", use_container_width=True)
            if submit:
                try:
                    response = client().auth.sign_in_with_password({"email": email, "password": password})
                    st.session_state.user = response.user
                    st.rerun()
                except Exception as error:
                    st.error(str(error))
        with tab_create:
            with st.form("signup"):
                email = st.text_input("Email address", key="new_email", placeholder="you@example.com")
                password = st.text_input("Password", key="new_password", type="password", help="At least 6 characters.")
                submit = st.form_submit_button("Create account", use_container_width=True)
            if submit:
                try:
                    response = client().auth.sign_up({"email": email, "password": password})
                    if response.user and response.session:
                        st.session_state.user = response.user
                        st.rerun()
                    st.success("Check your email to confirm your account, then sign in.")
                except Exception as error:
                    st.error(str(error))


def app_view(user):
    today = date.today()
    selected_month = st.session_state.get("month", today.month)
    selected_year = st.session_state.get("year", today.year)

    st.markdown(
        '<div class="topbar"><div class="brand-lockup"><div class="brand-mark">◈</div><span>HallBook</span></div><div class="topbar-copy"><span></span></div></div>',
        unsafe_allow_html=True)
    st.sidebar.markdown("### Account")
    st.sidebar.caption(user.email)
    st.sidebar.button("Sign out", on_click=sign_out, use_container_width=True)

    title_col, action_col = st.columns([4, 1])
    with title_col:
        st.markdown(
            "<p class='kicker'>HALL AVAILABILITY</p><h1>Plan something memorable.</h1><p class='subtle'>Booked time slots are visible to everyone. Your reservations are yours to cancel.</p>",
            unsafe_allow_html=True)
    with action_col:
        st.markdown("<div class='top-spacer'></div>", unsafe_allow_html=True)
        if st.button("＋ New booking", type="primary", use_container_width=True):
            st.session_state.open_booking = True

    if st.session_state.get("open_booking", False):
        booking_dialog(user)

    booking_data = get_bookings(selected_year, selected_month)

    # --- ADD THE DOWNLOAD BUTTON HERE ---
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Export Tools")
    download_excel_button(booking_data, selected_year, selected_month)
    # -------------------------------------

    # Month navigation UI code continues below...
    # c1, c2, c3 = st.columns()
    # Month navigation UI
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1:
        if st.button("◀ Previous", use_container_width=True):
            if selected_month == 1:
                st.session_state.month = 12
                st.session_state.year = selected_year - 1
            else:
                st.session_state.month = selected_month - 1
            st.rerun()
    with c2:
        st.markdown(
            f"<h3 style='text-align: center; margin-top: 0;'>{calendar.month_name[selected_month]} {selected_year}</h3>",
            unsafe_allow_html=True)
    with c3:
        if st.button("Next ▶", use_container_width=True):
            if selected_month == 12:
                st.session_state.month = 1
                st.session_state.year = selected_year + 1
            else:
                st.session_state.month = selected_month + 1
            st.rerun()

    # Render updated calendar grid
    html = calendar_html(selected_year, selected_month, booking_data)
    st.markdown(html, unsafe_allow_html=True)

    # Manage existing bookings display list below the calendar
    if booking_data:
        st.markdown("<br>### Booking Details", unsafe_allow_html=True)
        for b in booking_data:
            my_booking = b["booked_by"] == user.id
            col_details, col_cancel = st.columns([4, 1])
            with col_details:
                st.write(
                    f"**{b['title']}** | 📅 {format_day(b['booking_date_key'])} 🕒 {b['start_time_str']} - {b['end_time_str']} | 📞 {b['contact_number']}")
                if b["notes"]:
                    st.caption(f"Notes: {b['notes']}")
            with col_cancel:
                if my_booking:
                    if st.button("Cancel", key=f"del-{b['id']}", type="secondary", use_container_width=True):
                        try:
                            client().table("bookings").delete().eq("id", b["id"]).execute()
                            st.rerun()
                        except Exception as e:
                            st.error(str(e))


# Main routing architecture
if __name__ == "__main__":
    u = current_user()
    if u:
        app_view(u)
    else:
        login_view()
