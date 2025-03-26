import streamlit as st
import sqlite3
import pandas as pd
import re  # For phone number validation

st.set_page_config(page_title="Kanak Splitwise", page_icon="💰", layout="wide")

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
        share_per_person REAL
    )
""")
conn.commit()

# **Session state initialization**
if "logged_in_user" not in st.session_state:
    st.session_state["logged_in_user"] = None

# **Login Page (Shown if user is not logged in)**
if not st.session_state["logged_in_user"]:
    st.title("🔑 Kanak Splitwise Login")

    phone = st.text_input("📞 Enter your phone number").strip()
    full_name = st.text_input("👤 Enter your full name").strip()

    # Validate phone number
    def is_valid_phone(phone):
        return bool(re.fullmatch(r"^[6-9]\d{9}$", phone))

    if st.button("🔓 Login / Register"):
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
                    st.rerun()
                else:
                    st.error("❌ Incorrect name for this phone number!")
            else:
                cursor.execute("INSERT INTO users (phone, full_name) VALUES (?, ?)", (phone, full_name))
                conn.commit()
                st.session_state["logged_in_user"] = {"phone": phone, "name": full_name}
                st.rerun()
    
    st.stop()  # Stops execution if user is not logged in

# **Main Expense Tracker (Only visible after login)**
user_name = st.session_state["logged_in_user"]["name"]

st.title(f"💰 Kanak Splitwise - Welcome, {user_name}")

# Fetch users from database
cursor.execute("SELECT full_name FROM users")
all_users = [row[0] for row in cursor.fetchall()]

# **Expense Input Form**
st.subheader("📝 Add Expense")
with st.form("expense_form"):
    expense_desc = st.text_input("📌 Expense Description")
    amount = st.number_input("💵 Amount", min_value=0.0, format="%.2f")
    selected_people = st.multiselect("👥 Select people involved", all_users)

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

# **Display Expenses (Only for Logged-in User)**
st.subheader("📊 Your Expenses")
cursor.execute("SELECT * FROM expenses WHERE payer=? OR people LIKE ?", (user_name, f"%{user_name}%"))
expenses_data = cursor.fetchall()

if expenses_data:
    expense_df = pd.DataFrame(expenses_data, columns=["ID", "Payer", "Description", "Amount", "People", "Share Per Person"])

    # **Style names in "People" column with capsule effect**
    def format_people(people):
        names = people.split(",")
        return "".join([f'<span class="capsule">{name}</span> ' for name in names])

    expense_df["People"] = expense_df["People"].apply(format_people)
    expense_df.drop(columns=["ID"], inplace=True)

    # **Custom Styling**
    st.markdown(
        """
        <style>
        .capsule {
            background-color: #DDEEFF;
            color: #004488;
            padding: 5px 10px;
            border-radius: 15px;
            margin: 2px;
            display: inline-block;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # **Display styled dataframe**
    st.write(expense_df.to_html(escape=False), unsafe_allow_html=True)

# **Show Only the Logged-in User's Balance**
st.subheader("📉 Your Balance Summary")

balances = {}

for _, payer, _, amount, people, share_per_person in expenses_data:
    people_list = people.split(",")

    if payer == user_name:
        for person in people_list:
            if person != user_name:
                balances[person] = balances.get(person, 0) + share_per_person

    elif user_name in people_list:
        balances[payer] = balances.get(payer, 0) - share_per_person

# **Show summarized transactions per person**
for person, balance in balances.items():
    if balance > 0:
        st.write(f"🟢 {person} owes you ₹{balance:.2f}")
    elif balance < 0:
        st.write(f"🔴 You owe {person} ₹{-balance:.2f}")

# **Show final net balance**
total_balance = sum(balances.values())

st.subheader("⚖ Your Final Balance")
if total_balance > 0:
    st.write(f"🟢 You are owed ₹{total_balance:.2f}")
elif total_balance < 0:
    st.write(f"🔴 You owe ₹{-total_balance:.2f}")
else:
    st.write(f"⚖ You are settled")

# **Logout Button**
if st.button("🚪 Logout"):
    st.session_state["logged_in_user"] = None
    st.rerun()

# Close database connection
conn.close()
