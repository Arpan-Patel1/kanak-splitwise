import streamlit as st
import sqlite3
import pandas as pd
import re  # For phone number validation

st.set_page_config(page_title="Kanak Splitwise", page_icon="💰", layout="centered")

# Create SQLite connection
conn = sqlite3.connect("splitwise.db", check_same_thread=False)
cursor = conn.cursor()

# Create user table if not exists
cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        phone TEXT PRIMARY KEY,
        full_name TEXT
    )
""")

cursor.execute("""
    CREATE TABLE IF NOT EXISTS expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        payer TEXT,
        description TEXT,
        amount REAL,
        people TEXT,
        share_per_person REAL,
        settled INTEGER DEFAULT 0,
        settle_requested_by TEXT DEFAULT NULL
    )
""")
conn.commit()

# **Session state initialization**
if "logged_in_user" not in st.session_state:
    st.session_state["logged_in_user"] = None

# **Login Page (Shown if user is not logged in)**
if not st.session_state["logged_in_user"]:
    st.title("🔑 Kanak Splitwise Login")

    phone = st.text_input("📞 Enter Phone Number", placeholder="Enter 10-digit number").strip()
    full_name = st.text_input("👤 Full Name", placeholder="Enter your name").strip()

    # Validate phone number
    def is_valid_phone(phone):
        return bool(re.fullmatch(r"^[6-9]\d{9}$", phone))

    if st.button("Login / Register", use_container_width=True):
        if not is_valid_phone(phone):
            st.error("❌ Enter a valid 10-digit phone number starting with 6-9!")
        elif full_name:
            cursor.execute("SELECT full_name FROM users WHERE phone=?", (phone,))
            existing_user = cursor.fetchone()

            if existing_user:
                stored_name = existing_user[0].strip().lower()
                entered_name = full_name.strip().lower()

                if stored_name == entered_name:
                    st.session_state["logged_in_user"] = {"phone": phone, "name": existing_user[0]}
                    st.experimental_set_query_params(logged_in="true")
                    st.rerun()
                else:
                    st.error("❌ Incorrect name for this phone number!")
            else:
                cursor.execute("INSERT INTO users (phone, full_name) VALUES (?, ?)", (phone, full_name))
                conn.commit()
                st.session_state["logged_in_user"] = {"phone": phone, "name": full_name}
                st.experimental_set_query_params(logged_in="true")
                st.rerun()

    st.stop()  # Stops execution if user is not logged in

# **Main Expense Tracker (Only visible after login)**
user_name = st.session_state["logged_in_user"]["name"]

st.title(f"💰 Welcome, {user_name}")

# **Tabs for Different Sections**
tab1, tab2, tab3 = st.tabs(["➕ Add Expense", "📊 Balance & Tracking", "⏳ Pending Confirmations"])

# **Tab 1: Add Expense**
with tab1:
    st.subheader("📝 Add Expense")

    # Fetch users from database
    cursor.execute("SELECT full_name FROM users")
    all_users = [row[0] for row in cursor.fetchall()]

    with st.form("expense_form"):
        expense_desc = st.text_input("📌 Description", placeholder="E.g., Dinner, Cab Fare")
        amount = st.number_input("💵 Amount", min_value=0.0, format="%.2f")
        selected_people = st.multiselect("👥 Split With", all_users, default=[user_name])

        submit = st.form_submit_button("Add Expense")

        if submit:
            if expense_desc and amount > 0 and selected_people:
                share_per_person = round(amount / len(selected_people), 2)

                cursor.execute(
                    "INSERT INTO expenses (payer, description, amount, people, share_per_person) VALUES (?, ?, ?, ?, ?)",
                    (user_name, expense_desc, amount, ",".join(selected_people), share_per_person)
                )
                conn.commit()
                st.success("✅ Expense added successfully!")
                st.rerun()

# **Tab 2: Balance & Expense Tracking**
with tab2:
    st.subheader("📊 Your Expenses")

    cursor.execute("SELECT id, payer, description, amount, people, share_per_person, settled, settle_requested_by FROM expenses WHERE payer=? OR people LIKE ?", (user_name, f"%{user_name}%"))
    expenses_data = cursor.fetchall()

    if expenses_data:
        expense_df = pd.DataFrame(expenses_data, columns=["ID", "Payer", "Description", "Amount", "People", "Share Per Person", "Settled", "Settle Requested By"])
        
        # Convert "People" column into a readable format
        expense_df["People"] = expense_df["People"].apply(lambda x: ", ".join(x.split(",")))

        # **Display table with better readability**
        st.dataframe(expense_df.drop(columns=["ID", "Settled", "Settle Requested By"]), use_container_width=True)

    # **Balance Summary with Fixes**
    st.subheader("📉 Your Balance Summary")

    balances = {}
    expense_ids = {}
    settle_requests = {}

    for expense_id, payer, _, _, people, share_per_person, settled, settle_requested_by in expenses_data:
        if settled:
            continue

        people_list = people.split(",")

        if payer == user_name:
            for person in people_list:
                if person != user_name:
                    balances[person] = balances.get(person, 0) + share_per_person
                    expense_ids[(person, payer)] = expense_id
        elif user_name in people_list:
            balances[payer] = balances.get(payer, 0) - share_per_person
            expense_ids[(user_name, payer)] = expense_id

        if settle_requested_by:
            settle_requests[expense_id] = settle_requested_by

    if not balances:
        st.info("⚖ You're all settled!")
    else:
        for person, balance in balances.items():
            col1, col2 = st.columns([4, 1])

            with col1:
                if balance > 0:
                    st.success(f"🟢 {person} owes you ₹{balance:.2f}")
                else:
                    st.error(f"🔴 You owe {person} ₹{-balance:.2f}")

            with col2:
                expense_id = expense_ids.get((user_name, person))
                if expense_id and expense_id in settle_requests:
                    st.button("✅ Requested", key=f"requested_{person}", disabled=True)
                elif balance < 0 and st.button("Request Settle", key=f"request_settle_{person}"):
                    cursor.execute("UPDATE expenses SET settle_requested_by = ? WHERE id = ?", (user_name, expense_id))
                    conn.commit()
                    st.success(f"✅ Settlement request sent to {person}!")
                    st.rerun()

# **Tab 3: Pending Confirmations**
with tab3:
    st.subheader("⏳ Settlement Confirmations Needed")

    cursor.execute("SELECT id, payer, people, share_per_person, settle_requested_by FROM expenses WHERE payer=? AND settled=0 AND settle_requested_by IS NOT NULL", (user_name,))
    settle_requests = cursor.fetchall()

    if not settle_requests:
        st.info("✅ No pending settlements!")
    else:
        for expense_id, payer, people, share_per_person, requested_by in settle_requests:
            st.warning(f"📢 {requested_by} has requested to settle ₹{share_per_person:.2f}. Confirm settlement?")

            col1, col2 = st.columns(2)
            if col1.button("✅ Confirm", key=f"confirm_{expense_id}"):
                cursor.execute("UPDATE expenses SET settled = 1, settle_requested_by = NULL WHERE id = ?", (expense_id,))
                conn.commit()
                st.success(f"✅ Settlement with {requested_by} confirmed!")
                st.rerun()
            if col2.button("❌ Reject", key=f"reject_{expense_id}"):
                cursor.execute("UPDATE expenses SET settle_requested_by = NULL WHERE id = ?", (expense_id,))
                conn.commit()
                st.warning(f"❌ Settlement request from {requested_by} rejected!")
                st.rerun()

# **Logout Button**
if st.button("🚪 Logout", use_container_width=True):
    st.session_state["logged_in_user"] = None
    st.experimental_set_query_params(logged_in="false")
    st.rerun()

# Close database connection
conn.close()
