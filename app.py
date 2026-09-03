import datetime
import io
import re
import sqlite3
import pandas as pd
import streamlit as st

# =============================================================================
# 1. DATABASE ENGINE
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
    .stButton button { border-radius: 6px; font-weight: bold; }
    div[data-testid="stMetricValue"] { font-size: 22px; color: #2E67AE; }
    .email-text { font-size: 0.85em; color: #475569; word-break: break-all; }

    /* Group Color 0: Soft Sky Blue Container */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.group-marker-blue),
    div[data-testid="stBlock"]:has(.group-marker-blue),
    div[data-testid="element-container"]:has(.group-marker-blue) {
        background-color: #E0F2FE !important;
        border-left: 6px solid #0284C7 !important;
        border-radius: 8px !important;
        padding: 4px 8px !important;
        margin-bottom: 6px !important;
    }

    /* Group Color 1: Soft Peach / Orange Container */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.group-marker-orange),
    div[data-testid="stBlock"]:has(.group-marker-orange),
    div[data-testid="element-container"]:has(.group-marker-orange) {
        background-color: #FFEDD5 !important;
        border-left: 6px solid #EA580C !important;
        border-radius: 8px !important;
        padding: 4px 8px !important;
        margin-bottom: 6px !important;
    }

    /* Group Banner Badges */
    .banner-blue {
        background-color: #0284C7;
        color: #FFFFFF;
        padding: 4px 10px;
        border-radius: 4px;
        font-size: 0.82em;
        font-weight: bold;
        display: inline-block;
        margin-top: 10px;
        margin-bottom: 4px;
    }
    .banner-orange {
        background-color: #EA580C;
        color: #FFFFFF;
        padding: 4px 10px;
        border-radius: 4px;
        font-size: 0.82em;
        font-weight: bold;
        display: inline-block;
        margin-top: 10px;
        margin-bottom: 4px;
    }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# 3. HELPER FUNCTIONS & EXCEL CONVERTER
# =============================================================================
def format_display_group_name(raw_name):
    if not raw_name:
        return ""
    if "," in raw_name:
        parts = raw_name.split(",", 1)
        return f"{parts[1].strip()} {parts[0].strip()}"
    return raw_name.strip()

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Completed_CheckIns")
    return output.getvalue()

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

                buyer_full_display = f"{b_first} {b_last}".strip()
                primary_group_key = f"{b_last}, {b_first}".strip(", ") if b_last else buyer_full_display

                is_primary = 1 if (i == 0 and is_multi) else 0
                is_guest = 1 if i > 0 else 0

                if i == 0 and not f_name and not l_name:
                    f_name, l_name = b_first, b_last
                elif i > 0 and ((f_name.lower() == b_first.lower() and l_name.lower() == b_last.lower()) or (not f_name and not l_name)):
                    f_name = f"Guest {i}"
                    l_name = primary_group_key

                c_in_clean = "☑ YES" if c_in.upper() in ["YES", "Y", "TRUE", "1.0", "1", "CHECKED", "☑ YES", "☑", "✓ YES", "✓"] else "☐ NO"

                full1, full2 = f"{f_name} {l_name}".strip().lower(), f"{l_name} {f_name}".strip().lower()
                has_waiver = "☐ NO"
                if e_mail.lower() in signed_emails or full1 in signed_names or full2 in signed_names:
                    has_waiver = "☑ YES"

                rows.append((l_name, f_name, e_mail, t_desc, has_waiver, c_in_clean, t_in, primary_group_key, is_primary, is_guest))
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
            buyer_full_display = f"{f_name} {l_name}".strip()
            primary_group_key = f"{l_name}, {f_name}".strip(", ") if l_name else buyer_full_display

            full1, full2 = f_name.lower() + " " + l_name.lower(), l_name.lower() + " " + f_name.lower()
            has_waiver = "☐ NO"
            if e_mail.lower() in signed_emails or full1 in signed_names or full2 in signed_names:
                has_waiver = "☑ YES"

            rows.append((l_name, f_name, e_mail, t_desc, has_waiver, c_in_clean, t_in, primary_group_key, 1 if qty > 1 else 0, 0))

            for g in range(1, qty):
                guest_l_name = primary_group_key
                rows.append((guest_l_name, f"Guest {g}", e_mail, t_desc, "☐ NO", "☐ NO", "", primary_group_key, 0, 1))

    conn = get_db()
    conn.execute("DELETE FROM attendees")
    conn.executemany("""
        INSERT INTO attendees (last_name, first_name, email, ticket_type, completed_waiver, checked_in, time_checked_in, primary_name, is_primary, is_guest)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)
    conn.commit()
    conn.close()

# =============================================================================
# 4. HEADER & ROUTING PARAMETERS
# =============================================================================
url_line = st.query_params.get("line", "A-I")

st.markdown('<p class="main-header">I ❤ OAKLAND ALAMEDA ESTUARY</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Official Event Check-In & Waiver Management System (v6.4.5)</p>', unsafe_allow_html=True)

sidebar = st.sidebar
sidebar.header("⚙️ Event Control Panel")
app_mode = sidebar.radio("Select Interface Mode:", ["Volunteer Check-In Line", "Morning Admin Portal"])

LINE_GROUPS = {
    "A-I": ['A','B','C','D','E','F','G','H','I'],
    "J-R": ['J','K','L','M','N','O','P','Q','R'],
    "S-Z": ['S','T','U','V','W','X','Y','Z']
}

# =============================================================================
# 5. CARD RENDERER FUNCTION WITH CLASS MARKERS
# =============================================================================
def render_attendee_card(row, tab_prefix, color_idx):
    att_id = row["id"]
    is_checked = (row["checked_in"] == "☑ YES")
    has_waiver = (row["completed_waiver"] == "☑ YES")
    is_bold_primary = (row["is_primary"] == 1)

    marker_class = "group-marker-blue" if color_idx == 0 else "group-marker-orange"

    with st.container(border=True):
        st.markdown(f'<div class="{marker_class}"></div>', unsafe_allow_html=True)
        c1, c2, c3, c4, c5 = st.columns([2.5, 2.5, 1.2, 1.2, 1.2])
        
        name_str = f"**{row['first_name']} {row['last_name']}**" if is_bold_primary else f"{row['first_name']} {row['last_name']}"
        c1.markdown(f"{name_str}<br><span class='email-text'>{row['email']}</span>", unsafe_allow_html=True)
        c2.markdown(f"**Ticket:** {row['ticket_type']}")

        waiver_toggle = c3.checkbox("Waiver", value=has_waiver, key=f"w_{att_id}_{tab_prefix}")
        if waiver_toggle != has_waiver:
            conn = get_db()
            new_val = "☑ YES" if waiver_toggle else "☐ NO"
            conn.execute("UPDATE attendees SET completed_waiver = ? WHERE id = ?", (new_val, att_id))
            conn.commit()
            conn.close()
            st.rerun()

        checkin_toggle = c4.checkbox("Check-In", value=is_checked, key=f"c_{att_id}_{tab_prefix}")
        if checkin_toggle != is_checked:
            conn = get_db()
            new_val = "☑ YES" if checkin_toggle else "☐ NO"
            now_time = datetime.datetime.now().strftime("%I:%M:%S %p") if checkin_toggle else ""
            conn.execute("UPDATE attendees SET checked_in = ?, time_checked_in = ? WHERE id = ?", (new_val, now_time, att_id))
            conn.commit()
            conn.close()
            st.rerun()

        if c5.button("⚡ Both", key=f"b_{att_id}_{tab_prefix}"):
            conn = get_db()
            now_time = datetime.datetime.now().strftime("%I:%M:%S %p")
            conn.execute("UPDATE attendees SET completed_waiver = '☑ YES', checked_in = '☑ YES', time_checked_in = ? WHERE id = ?", (now_time, att_id))
            conn.commit()
            conn.close()
            st.rerun()

        with st.expander(f"✏️ Edit Info for {row['first_name']} {row['last_name']}"):
            with st.form(key=f"edit_form_{att_id}_{tab_prefix}"):
                ef1, ef2 = st.columns(2)
                new_first = ef1.text_input("First Name", value=row["first_name"])
                new_last = ef2.text_input("Last Name", value=row["last_name"])
                
                ef3, ef4 = st.columns(2)
                new_email = ef3.text_input("Email Address", value=row["email"])
                new_ticket = ef4.text_input("Ticket Type", value=row["ticket_type"])
                
                if st.form_submit_button("💾 Save Registrant Updates"):
                    conn = get_db()
                    conn.execute("""
                        UPDATE attendees 
                        SET first_name = ?, last_name = ?, email = ?, ticket_type = ? 
                        WHERE id = ?
                    """, (new_first.strip(), new_last.strip(), new_email.strip(), new_ticket.strip(), att_id))
                    conn.commit()
                    conn.close()
                    st.success("Registrant details updated successfully!")
                    st.rerun()

# =============================================================================
# 6. VOLUNTEER CHECK-IN INTERFACE
# =============================================================================
if app_mode == "Volunteer Check-In Line":
    st.subheader("📋 Volunteer Check-In Station")
    
    col_line, col_search, col_save = st.columns([1.2, 2, 1])
    with col_line:
        selected_line = st.selectbox(
            "Assigned Line Section:",
            ["A-I", "J-R", "S-Z", "All Lines"],
            index=["A-I", "J-R", "S-Z"].index(url_line) if url_line in ["A-I", "J-R", "S-Z"] else 0
        )
    with col_search:
        search_query = st.text_input("🔍 Search Registrant or Guest Name / Email:", "")
    with col_save:
        st.write(" ")
        st.write(" ")
        if st.button("💾 Save Progress", type="primary", use_container_width=True):
            st.toast(f"✅ Line {selected_line} progress synced & saved!", icon="💾")

    conn = get_db()
    
    counts_df = pd.read_sql_query("""
        SELECT LOWER(TRIM(primary_name)) as g_key, COUNT(*) as party_size 
        FROM attendees 
        GROUP BY LOWER(TRIM(primary_name))
    """, conn)
    party_sizes = dict(zip(counts_df['g_key'], counts_df['party_size']))

    query = "SELECT id, last_name, first_name, email, ticket_type, completed_waiver, checked_in, time_checked_in, primary_name, is_primary FROM attendees"
    params = []
    conditions = []
    
    if selected_line in LINE_GROUPS:
        letters = LINE_GROUPS[selected_line]
        letter_placeholders = ",".join(["?"] * len(letters))
        conditions.append(f"(UPPER(SUBSTR(last_name, 1, 1)) IN ({letter_placeholders}) OR UPPER(SUBSTR(primary_name, 1, 1)) IN ({letter_placeholders}))")
        params.extend(letters)
        params.extend(letters)

    if search_query.strip():
        sq = f"%{search_query.strip().lower()}%"
        conditions.append("(LOWER(last_name) LIKE ? OR LOWER(first_name) LIKE ? OR LOWER(email) LIKE ? OR LOWER(primary_name) LIKE ?)")
        params.extend([sq, sq, sq, sq])

    if conditions:
        query += " WHERE " + " AND ".join(conditions)

    query += " ORDER BY primary_name ASC, is_primary DESC, id ASC"

    df_display = pd.read_sql_query(query, conn, params=params)
    
    total_count = pd.read_sql_query("SELECT COUNT(*) as cnt FROM attendees", conn)["cnt"][0]
    checked_count = pd.read_sql_query("SELECT COUNT(*) as cnt FROM attendees WHERE checked_in = '☑ YES'", conn)["cnt"][0]
    conn.close()

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Event Registrations", total_count)
    m2.metric("Checked In Overall", checked_count)
    m3.metric("Section Roster Count", len(df_display))

    st.divider()

    tab_not_checked, tab_checked = st.tabs(["⏳ Pending Check-In", "✅ Checked In"])

    with tab_not_checked:
        df_pending = df_display[df_display["checked_in"] != "☑ YES"]
        if df_pending.empty:
            st.info("No pending attendees in this section.")
        else:
            current_group = None
            color_toggle = 0
            for idx, row in df_pending.iterrows():
                p_raw = str(row["primary_name"]).strip() if row["primary_name"] else f"{row['last_name']}, {row['first_name']}".strip(", ")
                g_key = p_raw.lower()
                party_size = party_sizes.get(g_key, 1)

                if g_key != current_group:
                    if current_group is not None:
                        color_toggle = 1 - color_toggle
                    current_group = g_key
                    
                    if party_size > 1:
                        banner_style = "banner-blue" if color_toggle == 0 else "banner-orange"
                        disp_name = format_display_group_name(p_raw)
                        st.markdown(f'<span class="{banner_style}">Group: {disp_name} ({party_size} People)</span>', unsafe_allow_html=True)

                render_attendee_card(row, tab_prefix="p", color_idx=color_toggle)

    with tab_checked:
        df_done = df_display[df_display["checked_in"] == "☑ YES"]
        if df_done.empty:
            st.info("No checked-in attendees in this section yet.")
        else:
            current_group = None
            color_toggle = 0
            for idx, row in df_done.iterrows():
                p_raw = str(row["primary_name"]).strip() if row["primary_name"] else f"{row['last_name']}, {row['first_name']}".strip(", ")
                g_key = p_raw.lower()
                party_size = party_sizes.get(g_key, 1)

                if g_key != current_group:
                    if current_group is not None:
                        color_toggle = 1 - color_toggle
                    current_group = g_key
                    
                    if party_size > 1:
                        banner_style = "banner-blue" if color_toggle == 0 else "banner-orange"
                        disp_name = format_display_group_name(p_raw)
                        st.markdown(f'<span class="{banner_style}">Group: {disp_name} ({party_size} People)</span>', unsafe_allow_html=True)

                render_attendee_card(row, tab_prefix="c", color_idx=color_toggle)

# =============================================================================
# 7. MORNING ADMIN PORTAL & SECTION EXPORTS
# =============================================================================
elif app_mode == "Morning Admin Portal":
    st.subheader("🔑 Morning Admin Portal & Master File Management")
    admin_pass = sidebar.text_input("Enter Admin Password:", type="password")
    
    if admin_pass != "ihoae2026":
        st.warning("Please enter the Admin Password in the sidebar to access management settings.")
    else:
        st.success("Admin Authenticated!")

        st.markdown("### 🗑️ System Reset")
        if st.button("Clear All Database Records", type="secondary"):
            conn = get_db()
            conn.execute("DELETE FROM attendees")
            conn.commit()
            conn.close()
            st.success("Database cleared! All records reset to 0.")
            st.rerun()

        st.divider()

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
                st.success("Event database loaded successfully!")
                st.rerun()

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
                    p_key = f"{w_last}, {w_first}".strip(", ")
                    conn.execute("""
                        INSERT INTO attendees (last_name, first_name, email, ticket_type, completed_waiver, checked_in, time_checked_in, primary_name, is_primary, is_guest)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0, 0)
                    """, (w_last, w_first, w_email, w_ticket, "☑ YES" if w_waiver else "☐ NO", "☑ YES" if w_checkin else "☐ NO", now_str, p_key))
                    conn.commit()
                    conn.close()
                    st.success(f"Added walk-in registration for {w_first} {w_last}!")
                    st.rerun()

        st.divider()

        st.markdown("### 3. Section-Specific Excel Exports")
        st.info("Download individual Excel spreadsheets (.xlsx) containing attendees who have completed BOTH Waiver & Check-In.")

        conn = get_db()
        
        tab_sec1, tab_sec2, tab_sec3, tab_master = st.tabs([
            "Line 1: Section A–I", 
            "Line 2: Section J–R", 
            "Line 3: Section S–Z", 
            "All Attendees (Master)"
        ])

        def get_section_query(letters):
            placeholders = ",".join(["?"] * len(letters))
            q = f"""
                SELECT 
                    last_name AS [Last Name], 
                    first_name AS [First Name], 
                    email AS [Email], 
                    ticket_type AS [Ticket Type], 
                    completed_waiver AS [Signed Waiver?], 
                    checked_in AS [Checked In?], 
                    time_checked_in AS [Check-In Time],
                    primary_name AS [Primary Signee Key]
                FROM attendees 
                WHERE (UPPER(SUBSTR(last_name, 1, 1)) IN ({placeholders}) 
                   OR UPPER(SUBSTR(primary_name, 1, 1)) IN ({placeholders}))
                  AND completed_waiver = '☑ YES'
                  AND checked_in = '☑ YES'
                ORDER BY primary_name ASC, is_primary DESC, id ASC
            """
            return pd.read_sql_query(q, conn, params=letters + letters)

        with tab_sec1:
            df_sec1 = get_section_query(LINE_GROUPS["A-I"])
            st.dataframe(df_sec1, use_container_width=True)
            excel_bytes1 = convert_df_to_excel(df_sec1)
            st.download_button(
                label="💾 Download Line 1 (A–I) Completed Excel Roster (.xlsx)",
                data=excel_bytes1,
                file_name=f"IHOAE_Section_A-I_CheckedIn_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        with tab_sec2:
            df_sec2 = get_section_query(LINE_GROUPS["J-R"])
            st.dataframe(df_sec2, use_container_width=True)
            excel_bytes2 = convert_df_to_excel(df_sec2)
            st.download_button(
                label="💾 Download Line 2 (J–R) Completed Excel Roster (.xlsx)",
                data=excel_bytes2,
                file_name=f"IHOAE_Section_J-R_CheckedIn_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        with tab_sec3:
            df_sec3 = get_section_query(LINE_GROUPS["S-Z"])
            st.dataframe(df_sec3, use_container_width=True)
            excel_bytes3 = convert_df_to_excel(df_sec3)
            st.download_button(
                label="💾 Download Line 3 (S–Z) Completed Excel Roster (.xlsx)",
                data=excel_bytes3,
                file_name=f"IHOAE_Section_S-Z_CheckedIn_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        with tab_master:
            df_master = pd.read_sql_query("""
                SELECT 
                    last_name AS [Last Name], 
                    first_name AS [First Name], 
                    email AS [Email], 
                    ticket_type AS [Ticket Type], 
                    completed_waiver AS [Signed Waiver?], 
                    checked_in AS [Checked In?], 
                    time_checked_in AS [Check-In Time], 
                    primary_name AS [Primary Signee Key] 
                FROM attendees 
                WHERE completed_waiver = '☑ YES' 
                  AND checked_in = '☑ YES'
                ORDER BY primary_name ASC, is_primary DESC, id ASC
            """, conn)
            st.dataframe(df_master, use_container_width=True)
            excel_bytes_master = convert_df_to_excel(df_master)
            st.download_button(
                label="💾 Download Complete Master Completed Excel Roster (.xlsx)",
                data=excel_bytes_master,
                file_name=f"IHOAE_Master_CheckedIn_Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        conn.close()
