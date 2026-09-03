import datetime
import re
import sqlite3
import pandas as pd
import streamlit as st

# =============================================================================
# 1. DATABASE & PERSISTENCE ENGINE
# =============================================================================
DB_FILE = "checkin_master.db"

def get_db():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS attendees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            last_name TEXT,
            first_name TEXT,
            email TEXT,
            ticket_type TEXT,
            completed_waiver TEXT,
            checked_in TEXT,
            time_checked_in TEXT,
            primary_name TEXT,
            is_primary INTEGER,
            is_guest INTEGER
        )
    """)
    conn.commit()
    conn.close()

init_db()

# =============================================================================
# 2. PAGE CONFIG & STYLING
# =============================================================================
st.set_page_config(
    page_title="IHOAE Event Check-In Manager",
    page_icon="🌊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header { font-size: 26px; font-weight: bold; color: #1B4173; margin-bottom: 0px; }
    .sub-header { font-size: 14px; font-style: italic; color: #FF6F43; margin-bottom: 15px; }
    .stButton button { border-radius: 8px; font-weight: bold; }
    div[data-testid="stMetricValue"] { font-size: 22px; color: #2E67AE; }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# 3. HELPER FUNCTIONS & WAIVER PARSING
# =============================================================================
def parse_waivers(file_obj):
    names, emails = set(), set()
    try:
        df = pd.read_csv(file_obj) if file_obj.name.endswith(".csv") else pd.read_excel(file_obj)
        content = df.to_csv(index=False, sep="\t")
        for em in re.findall(r"[a-zA-Z0-9%._+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", content):
            emails.add(em.strip().lower())
        for col in [c for c in df.columns if "name" in str(c).lower()]:
            for val in df[col].fillna("").astype(str):
                if val.strip() and val.lower() != "nan":
                    for sn in val.split(";"):
                        if len(sn.strip()) > 2:
                            names.add(sn.strip().lower())
    except Exception as e:
        st.error(f"Error reading waiver file: {e}")
    return names, emails

def populate_database(df_imported, signed_names, signed_emails):
    df_imported.columns = [str(c).replace("\t", "").strip() for c in df_imported.columns]
    lower_cols = {c.lower(): c for c in df_imported.columns}

    def find_col(cands):
        return next((lower_cols[c] for c in cands if c in lower_cols), None)

    first_col = find_col(["attendee first name", "first name", "firstname", "first"])
    last_col = find_col(["attendee last name", "last name", "lastname", "last"])
    buyer_first_col = find_col(["buyer first name", "purchaser first name"])
    buyer_last_col = find_col(["buyer last name", "purchaser last name"])
    full_name_col = find_col(["name", "full name", "attendee name", "buyer name"])
    tickets_qty_col = find_col(["tickets", "ticket quantity", "quantity", "qty"])
    ticket_type_col = find_col(["ticket type", "ticket tier", "activity type", "category"])
    event_name_col = find_col(["event name"])
    email_col = find_col(["email", "email address", "buyer email", "attendee email"])
    checkin_col = find_col(["checked in?", "checked in", "check in", "status"])
    time_col = find_col(["time checked in", "check in time", "timestamp"])
    order_id_col = find_col(["order id", "order #", "order number"])

    rows = []
    is_attendees_export = ("attendee first name" in lower_cols or "attendee last name" in lower_cols) and (order_id_col is not None)

    if is_attendees_export:
        grouped = df_imported.groupby(order_id_col, sort=False)
        for order_id, group in grouped:
            is_multi = len(group) > 1
            for i, (orig_idx, r) in enumerate(group.iterrows()):
                f_name = str(r[first_col]).strip() if first_col and pd.notna(r[first_col]) else ""
                l_name = str(r[last_col]).strip() if last_col and pd.notna(r[last_col]) else ""
                b_first = str(r[buyer_first_col]).strip() if buyer_first_col and pd.notna(r[buyer_first_col]) else f_name
                b_last = str(r[buyer_last_col]).strip() if buyer_last_col and pd.notna(r[buyer_last_col]) else l_name
                e_mail = str(r[email_col]).strip() if email_col and pd.notna(r[email_col]) else ""

                t_desc = "General"
                if ticket_type_col and pd.notna(r[ticket_type_col]):
                    val = str(r[ticket_type_col]).strip()
                    if val.lower() not in ["free order", "paid order", "order", "1", "2", "3", "4"]:
                        t_desc = val
                    elif event_name_col and pd.notna(r[event_name_col]):
                        t_desc = str(r[event_name_col]).strip()
                elif event_name_col and pd.notna(r[event_name_col]):
                    t_desc = str(r[event_name_col]).strip()

                c_in = str(r[checkin_col]).strip() if checkin_col and pd.notna(r[checkin_col]) else "☐ NO"
                t_in = str(r[time_col]).strip() if time_col and pd.notna(r[time_col]) else ""

                buyer_full = f"{b_first} {b_last}".strip()
                is_primary = 1 if (i == 0 and is_multi) else 0
                is_guest = 1 if i > 0 else 0

                if i == 0 and not f_name and not l_name:
                    f_name, l_name = b_first, b_last
                elif i > 0 and ((f_name.lower() == b_first.lower() and l_name.lower() == b_last.lower()) or (not f_name and not l_name)):
                    f_name = f"Guest {i}"
                    l_name = f"({buyer_full})" if buyer_full else ""

                c_in_clean = "☑ YES" if c_in.upper() in ["YES", "Y", "TRUE", "1.0", "1", "CHECKED", "☑ YES", "☑", "✓ YES", "✓"] else "☐ NO"

                # Check waiver
                full1, full2 = f"{f_name} {l_name}".strip().lower(), f"{l_name} {f_name}".strip().lower()
                has_waiver = "☐ NO"
                if e_mail.lower() in signed_emails or full1 in signed_names or full2 in signed_names:
                    has_waiver = "☑ YES"

                rows.append((l_name, f_name, e_mail, t_desc, has_waiver, c_in_clean, t_in, buyer_full, is_primary, is_guest))
    else:
        for orig_idx, r in df_imported.iterrows():
            f_name = str(r[first_col]).strip() if first_col and pd.notna(r[first_col]) else ""
            l_name = str(r[last_col]).strip() if last_col and pd.notna(r[last_col]) else ""
            if not f_name and not l_name and full_name_col and pd.notna(r[full_name_col]):
                parts = str(r[full_name_col]).strip().split()
                f_name = parts[0] if len(parts) > 0 else ""
                l_name = " ".join(parts[1:]) if len(parts) > 1 else ""

            e_mail = str(r[email_col]).strip() if email_col and pd.notna(r[email_col]) else ""

            t_desc = "General"
            if ticket_type_col and pd.notna(r[ticket_type_col]):
                val = str(r[ticket_type_col]).strip()
                if val.lower() not in ["free order", "paid order", "order", "1", "2", "3", "4"]:
                    t_desc = val
                elif event_name_col and pd.notna(r[event_name_col]):
                    t_desc = str(r[event_name_col]).strip()
            elif event_name_col and pd.notna(r[event_name_col]):
                t_desc = str(r[event_name_col]).strip()

            c_in = str(r[checkin_col]).strip() if checkin_col and pd.notna(r[checkin_col]) else "☐ NO"
            t_in = str(r[time_col]).strip() if time_col and pd.notna(r[time_col]) else ""

            qty = 1
            if tickets_qty_col and pd.notna(r[tickets_qty_col]):
                try:
                    qty = int(r[tickets_qty_col])
                except ValueError:
                    qty = 1

            c_in_clean = "☑ YES" if c_in.upper() in ["YES", "Y", "TRUE", "1.0", "1", "CHECKED", "☑ YES", "☑", "✓ YES", "✓"] else "☐ NO"
            buyer_full = f"{f_name} {l_name}".strip()

            full1, full2 = f_name.lower() + " " + l_name.lower(), l_name.lower() + " " + f_name.lower()
            has_waiver = "☐ NO"
            if e_mail.lower() in signed_emails or full1 in signed_names or full2 in signed_names:
                has_waiver = "☑ YES"

            rows.append((l_name, f_name, e_mail, t_desc, has_waiver, c_in_clean, t_in, buyer_full, 1 if qty > 1 else 0, 0))

            for g in range(1, qty):
                rows.append((f"({buyer_full})" if buyer_full else "", f"Guest {g}", e_mail, t_desc, "☐ NO", "☐ NO", "", buyer_full, 0, 1))

    conn = get_db()
    conn.execute("DELETE FROM attendees")
    conn.executemany("""
        INSERT INTO attendees (last_name, first_name, email, ticket_type, completed_waiver, checked_in, time_checked_in, primary_name, is_primary, is_guest)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)
    conn.commit()
    conn.close()

# =============================================================================
# 4. ROUTING & QUERY PARAMETERS
# =============================================================================
# Read line selection from URL parameter (default to Line 1: A-I)
url_line = st.query_params.get("line", "A-I")

st.markdown('<p class="main-header">I ❤ OAKLAND ALAMEDA ESTUARY</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Official Event Check-In & Waiver Management System (v6.1)</p>', unsafe_allow_html=True)

sidebar = st.sidebar
sidebar.header("⚙️ Event Control Panel")
app_mode = sidebar.radio("Select Interface Mode:", ["Volunteer Check-In Line", "Morning Admin Portal"])

# Alphabetical Line Groupings
LINE_GROUPS = {
    "A-I": ['A','B','C','D','E','F','G','H','I'],
    "J-R": ['J','K','L','M','N','O','P','Q','R'],
    "S-Z": ['S','T','U','V','W','X','Y','Z']
}

# =============================================================================
# 5. VOLUNTEER CHECK-IN INTERFACE
# =============================================================================
if app_mode == "Volunteer Check-In Line":
    st.subheader("📋 Volunteer Check-In Station")
    
    col_line, col_search = st.columns([1, 2])
    with col_line:
        selected_line = st.selectbox(
            "Assigned Line Section:",
            ["A-I", "J-R", "S-Z", "All Lines"],
            index=["A-I", "J-R", "S-Z"].index(url_line) if url_line in ["A-I", "J-R", "S-Z"] else 0
        )
    with col_search:
        search_query = st.text_input("🔍 Search Registrant or Guest Name / Email:", "")

    conn = get_db()
    
    # Base query
    query = "SELECT id, last_name, first_name, email, ticket_type, completed_waiver, checked_in, time_checked_in, primary_name, is_primary FROM attendees"
    params = []

    conditions = []
    
    # Filter by Alphabet Section
    if selected_line in LINE_GROUPS:
        letters = LINE_GROUPS[selected_line]
        letter_placeholders = ",".join(["?"] * len(letters))
        # Match primary buyer's last name initial OR guest's last name initial
        conditions.append(f"(UPPER(SUBSTR(last_name, 1, 1)) IN ({letter_placeholders}) OR UPPER(SUBSTR(primary_name, 1, 1)) IN ({letter_placeholders}))")
        params.extend(letters)
        params.extend(letters)

    # Filter by Search Query
    if search_query.strip():
        sq = f"%{search_query.strip().lower()}%"
        conditions.append("(LOWER(last_name) LIKE ? OR LOWER(first_name) LIKE ? OR LOWER(email) LIKE ? OR LOWER(primary_name) LIKE ?)")
        params.extend([sq, sq, sq, sq])

    if conditions:
        query += " WHERE " + " AND ".join(conditions)

    query += " ORDER BY primary_name ASC, is_primary DESC, id ASC"

    df_display = pd.read_sql_query(query, conn, params=params)
    conn.close()

    # Metrics Summary
    conn = get_db()
    total_count = pd.read_sql_query("SELECT COUNT(*) as cnt FROM attendees", conn)["cnt"][0]
    checked_count = pd.read_sql_query("SELECT COUNT(*) as cnt FROM attendees WHERE checked_in = '☑ YES'", conn)["cnt"][0]
    conn.close()

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Event Registrations", total_count)
    m2.metric("Checked In Overall", checked_count)
    m3.metric("Section Roster Count", len(df_display))

    st.divider()

    if df_display.empty:
        st.info("No attendees found for this section or search query.")
    else:
        for idx, row in df_display.iterrows():
            att_id = row["id"]
            is_checked = (row["checked_in"] == "☑ YES")
            has_waiver = (row["completed_waiver"] == "☑ YES")
            is_bold_primary = (row["is_primary"] == 1)

            # Styling container
            bg_color = "#EEF9F1" if is_checked else "#FFFFFF"
            border_color = "#16A34A" if is_checked else "#D0E1F0"

            with st.container():
                c1, c2, c3, c4, c5 = st.columns([2.5, 2.5, 1.5, 1.5, 1.5])
                
                # Name Display
                name_str = f"**{row['first_name']} {row['last_name']}**" if is_bold_primary else f"{row['first_name']} {row['last_name']}"
                c1.markdown(f"{name_str}\n\n`<small>{row['email']}</small>`", unsafe_allow_html=True)
                c2.markdown(f"**Ticket:** {row['ticket_type']}")

                # Waiver Checkbox
                waiver_toggle = c3.checkbox("Waiver", value=has_waiver, key=f"w_{att_id}")
                if waiver_toggle != has_waiver:
                    conn = get_db()
                    new_val = "☑ YES" if waiver_toggle else "☐ NO"
                    conn.execute("UPDATE attendees SET completed_waiver = ? WHERE id = ?", (new_val, att_id))
                    conn.commit()
                    conn.close()
                    st.rerun()

                # Check-In Checkbox
                checkin_toggle = c4.checkbox("Check-In", value=is_checked, key=f"c_{att_id}")
                if checkin_toggle != is_checked:
                    conn = get_db()
                    new_val = "☑ YES" if checkin_toggle else "☐ NO"
                    now_time = datetime.datetime.now().strftime("%I:%M:%S %p") if checkin_toggle else ""
                    conn.execute("UPDATE attendees SET checked_in = ?, time_checked_in = ? WHERE id = ?", (new_val, now_time, att_id))
                    conn.commit()
                    conn.close()
                    st.rerun()

                # Check Both Button
                if c5.button("⚡ Both", key=f"b_{att_id}"):
                    conn = get_db()
                    now_time = datetime.datetime.now().strftime("%I:%M:%S %p")
                    conn.execute("UPDATE attendees SET completed_waiver = '☑ YES', checked_in = '☑ YES', time_checked_in = ? WHERE id = ?", (now_time, att_id))
                    conn.commit()
                    conn.close()
                    st.rerun()

            st.divider()

# =============================================================================
# 6. MORNING ADMIN PORTAL & MASTER EXPORT
# =============================================================================
elif app_mode == "Morning Admin Portal":
    st.subheader("🔑 Morning Admin Portal & Master File Management")
    admin_pass = sidebar.text_input("Enter Admin Password:", type="password")
    
    if admin_pass != "ihoae2026":
        st.warning("Please enter the Admin Password in the sidebar to access management settings.")
    else:
        st.success("Admin Authenticated!")
        
        st.markdown("### 1. Pre-Load Morning Event Files")
        st.info("Upload your Eventbrite export and Waiver file here. This will populate the master check-in list for all volunteer phones.")

        file_orders = st.file_uploader("Upload Eventbrite Orders / Attendees List (.xls, .xlsx, .csv)", type=["xls", "xlsx", "csv"])
        file_waivers = st.file_uploader("Upload Signed Waiver List (.csv, .xlsx)", type=["csv", "xlsx", "xls"])

        if st.button("🚀 Process & Pre-Load Event Data", type="primary"):
            if not file_orders:
                st.error("Please upload at least the Eventbrite Orders or Attendees file.")
            else:
                s_names, s_emails = parse_waivers(file_waivers) if file_waivers else (set(), set())
                try:
                    df_imp = pd.read_csv(file_orders) if file_orders.name.endswith(".csv") else pd.read_excel(file_orders)
                except Exception:
                    df_imp = pd.read_csv(file_orders, sep="\t")
                
                populate_database(df_imp, s_names, s_emails)
                st.success("Event database loaded successfully! Volunteer phones can now access the list via their assigned line links.")

        st.divider()

        st.markdown("### 2. Add On-the-Fly Walk-In Registration")
        with st.form("walkin_form", clear_on_submit=True):
            w_first = st.text_input("First Name:")
            w_last = st.text_input("Last Name:")
            w_email = st.text_input("Email Address:")
            w_ticket = st.selectbox("Ticket Type:", ["Walk-In Cleanup", "General Registration", "Kayak / Canoe / SUP", "Walking Shore Cleanup"])
            w_waiver = st.checkbox("Signed Waiver?", value=True)
            w_checkin = st.checkbox("Check In Immediately?", value=True)

            if st.form_submit_button("➕ Add Walk-In Attendee"):
                if not w_first and not w_last:
                    st.error("Please enter a name for the walk-in attendee.")
                else:
                    now_str = datetime.datetime.now().strftime("%I:%M:%S %p") if w_checkin and w_waiver else ""
                    conn = get_db()
                    conn.execute("""
                        INSERT INTO attendees (last_name, first_name, email, ticket_type, completed_waiver, checked_in, time_checked_in, primary_name, is_primary, is_guest)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0, 0)
                    """, (w_last, w_first, w_email, w_ticket, "☑ YES" if w_waiver else "☐ NO", "☑ YES" if w_checkin else "☐ NO", now_str, f"{w_first} {w_last}".strip()))
                    conn.commit()
                    conn.close()
                    st.success(f"Added walk-in registration for {w_first} {w_last}!")

        st.divider()

        st.markdown("### 3. Master Roster Export")
        st.info("Retrieve the final consolidated check-in list with all check-in timestamps and walk-in additions.")

        conn = get_db()
        df_master = pd.read_sql_query("SELECT last_name AS [Last Name], first_name AS [First Name], email AS [Email], ticket_type AS [Ticket Type], completed_waiver AS [Signed Waiver?], checked_in AS [Checked In?], time_checked_in AS [Check-In Time] FROM attendees", conn)
        conn.close()

        st.dataframe(df_master, use_container_width=True)

        csv_bytes = df_master.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="💾 Download Consolidated Master Check-In Roster (CSV)",
            data=csv_bytes,
            file_name=f"IHOAE_Master_CheckIn_Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )