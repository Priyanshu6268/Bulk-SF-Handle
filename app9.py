import pandas as pd
import streamlit as st
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

# Function to send email
def send_email(subject, body, cc_addresses, smtp_server, smtp_port, sender_email, sender_password):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = 'priyanshukumar7470@gmail.com'
        msg['Subject'] = subject
        msg['Cc'] = ', '.join(cc_addresses)

        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            text = msg.as_string()
            server.sendmail(sender_email, [msg['To']] + cc_addresses, text)

        return f"Email sent to {msg['To']} with CC to {', '.join(cc_addresses)} and subject '{subject}'"
    except Exception as e:
        return f"Failed to send email: {str(e)}"

# Set the background style and logo
background_style = """
<style>
[data-testid="stAppViewContainer"] > .main {
    background-image: url("https://cdn.neowin.com/news/images/uploaded/2023/02/1677429005_aip3aiw-zvjzcnzj.gif");
    background-size: cover;
    background-position: center;
    color: #333; /* Default text color */
    padding: 20px;
    border-radius: 20px; /* Rounded corners */
}

.logo {
    position: absolute;
    top: 20px;
    left: -200px;
    width: 150px; /* Adjust size as needed */
}

h1, h2, h3, h4, h5, h6 {
    color: #004b90; /* Heading color */
    font-family: 'Helvetica Neue', sans-serif; /* Font family */
    font-size: 40px; /* Increased font size */
    font-weight: bold; /* Bold text */
}

.footer {
    text-align: center; 
    margin-top: 50px; 
    color: #004b87;
    font-size: 35px; /* Increased font size */
    font-weight: bold; /* Bold text */
}

input[type="text"], input[type="password"] {
    background-color: rgba(255, 255, 255, 0.9); /* Slightly transparent white */
    color: #000;  /* Change text color inside the input box */
    border: 2px solid #004b87; /* Border color */
    border-radius: 5px; /* Rounded corners */
    padding: 10px;
    font-size: 16px; /* Increased font size */
    font-weight: bold; /* Bold text */
}

input[type="text"]:focus, input[type="password"]:focus {
    border-color: #00f2fe; /* Border color on focus */
    outline: none;
    box-shadow: 0 0 5px rgba(0, 240, 254, 0.8); /* Shadow effect */
}

.stTextInput label {
    margin-bottom: -40px; /* Reduce margin below label */
}

.stTextInput {
    margin-bottom: 10px; /* Reduce margin below input */
}
</style>
"""

st.markdown(background_style, unsafe_allow_html=True)

# Add logo
st.markdown('<img class="logo" src="https://www.nokia.com/sites/default/files/2023-02/nokia-refreshed-logo-1_1.jpg?height=244&width=543" alt="Nokia Logo">', unsafe_allow_html=True)

# Centered title
st.markdown(
    """
    <h1 style='text-align: center; color: white; font-family: "Algerian"; font-size: 48px;'>Bulk SF-Case Update</h1>
    """,
    unsafe_allow_html=True
)

# SMTP details
smtp_server = "smtp.office365.com"
smtp_port = 587
sender_email = "priyanshukumarsaw@outlook.com"
sender_password = 'idulybvwptgeaiax'

st.markdown("<span style='font-size: 20px; font-weight: bold;'>Provide additional CC emails (comma-separated):</span>", unsafe_allow_html=True)
sender_cc = st.text_input("")

# File uploader with increased visibility
st.markdown('<div style="font-size: 20px; font-weight: bold; color: #004b87;">Choose an Excel file:</div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("Data from Excel file:")
    st.write(df)

    if st.button('Send Emails'):
        if not sender_email or not sender_password:
            st.error("Please provide email and password to send emails.")
        else:
            results = []
            # Split CC addresses by comma and strip whitespace
            cc_addresses = [email.strip() for email in sender_cc.split(',') if email.strip()]
            for _, row in df.iterrows():
                owner_mail = row['OwnerEmail'].strip()
                dt = datetime.now()
                update = f"Updated as on {dt:%I:%M:%S %p, %B %d, %Y}{'\n'}{'\n'}{row['Update']}"
                subject = row['Subject']
                body = update
                # Append owner mail to CC addresses
                cc_addresses.append(owner_mail)
                result = send_email(subject, body, cc_addresses, smtp_server, smtp_port, sender_email, sender_password)
                cc_addresses.pop()
                results.append(result)

            st.write("Email sending results:")
            for result in results:
                st.write(result)

# Copyright information
st.markdown(
    """
    <div class="footer">
        <p>&copy; 2024 Nokia. All rights reserved by MN S GSD SWSS NM&SON TS Team 4.</p>
    </div>
    """,
    unsafe_allow_html=True
)
