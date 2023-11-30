# Bulk_email_sender
Thought of this app idea during my application for my master's program to help me send more and yet still personalized email. Hopefully you find it helpful too.


# Bulk Email Sender

**Bulk Email Sender** is a Python application developed by Zikonde that provides a user-friendly interface for sending bulk emails. The program allows users to customize emails using parameters extracted from brackets within the email body. It supports attachments and saves sent emails to the "sent" folder on the email server.

## Features

- **User-Friendly Interface**: Built using the Tkinter library, the application provides a clean and intuitive user interface for easy navigation.

- **Dark Mode**: The application features a dark mode for a visually appealing experience.

- **Multi-Platform Support**: As a Python script, the application can run on various platforms, offering flexibility to users.

- **Parameterized Emails**: Users can customize emails by including parameters within curly brackets in the email body. The program dynamically replaces these parameters with corresponding values from an Excel file.

- **Attachment Support**: Users can attach files to their emails, enhancing the richness of their communication.

- **Sent Email Logging**: Sent emails are automatically saved to the "sent" folder on the email server, providing users with a record of their sent communications.

## Getting Started

1. **Clone the Repository**: Clone this repository to your local machine.

   ```bash
   git clone https://github.com/your-username/bulk-email-sender.git
   ```

2. **Install Dependencies**: Install the required Python libraries.

   ```bash
   pip install pandas openpyxl
   ```

3. **Run the Application**: Execute the Python script to launch the Bulk Email Sender application.

   ```bash
   python bulk_email_sender.py
   ```

4. **Configure Email Settings**: Enter your email, host, and password details. Select the appropriate host from the provided options (`gmail.com`, `qq.com`, `163.com`).

5. **Compose Email**: Fill in the sender email, subject, and body. Optionally, include parameters within curly brackets to personalize emails.

6. **Attach Files**: Attach files to your emails by clicking the "Attach File" button.

7. **Select Excel File**: Choose an Excel file containing recipient details.

8. **Send Emails**: Click the "Send Emails" button to start sending personalized emails.

## Note

- **SMTP and IMAP Configuration**: Ensure that your email provider's SMTP and IMAP settings are correctly configured for the application to send and save emails.

- **Security**: It's recommended to use an application-specific password instead of your main email password for increased security.

## Acknowledgments

This application utilizes various Python libraries, including Tkinter, Pandas, and Openpyxl. Special thanks to Zikonde for developing this useful tool.

**Disclaimer**: Ensure that your use of this application complies with the terms of service of your email provider. The application may require adjustments based on your specific email provider's settings.

Feel free to contribute, report issues, or suggest improvements by creating issues or pull requests on the GitHub repository.

Happy emailing! 