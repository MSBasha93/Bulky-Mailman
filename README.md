# Bulky-Mailman
Bulk email sender script to cold approach people with caution lol

## Overview
This application is a **web-based** tool for sending personalized bulk emails using a GoDaddy SMTP account. Users can access the application through a secure online portal, load a list of recipients from a CSV or Excel file, and send emails with dynamic fields like {name}, {company}, and {info}.

## Features
- Secure online access to the application.
- Load recipient lists from `.csv` or `.xlsx` files.
- Personalize email body using `{name}`, `{company}`, and `{info}` placeholders.
- Send emails via GoDaddy SMTP (`smtpout.secureserver.net`).
- Adjustable delay between emails.
- Kill button to stop email sending immediately and receive a summary report.
- Automatic sending of a success/failure summary report to the sender's email.
- Real-time progress bar showing sending status.
- Responsive and user-friendly web interface.

## File Format Requirements
The CSV/Excel file should contain at least the following columns:
- `name`: Recipient's Name
- `email`: Recipient's Email Address
- `company_name`: Name of the Company
- `company_info`: Short Description or Info about the Company

Example CSV:
```csv
name,email,company_name,company_info
John Doe,john@example.com,Example Inc.,Leading innovator in tech.
Jane Smith,jane@example.com,TechWorld,Specialized in AI solutions.
```

## Usage Instructions
1. Access the application via your web browser at [YourWebsite.com/email-sender](https://YourWebsite.com/email-sender).
2. Enter your GoDaddy email credentials securely.
3. Set the email subject and body. Use placeholders `{name}`, `{company}`, `{info}` where necessary.
4. Upload the recipients file (CSV/Excel).
5. Set delay between sending emails if needed.
6. Click **Send Emails**.
7. Monitor progress via the progress bar.
8. If needed, click **Kill Process** to immediately stop sending.
9. After completion, a summary report will be emailed to you.

## Important Notes
- Your credentials are encrypted and securely used only during your session.
- Use strong SMTP credentials.
- Killing the process mid-way will still generate a partial summary report.
- Ensure proper formatting of input CSV/Excel files to avoid runtime errors.
- For your security, logout after each session.

## Limitations
- Only tested with GoDaddy SMTP accounts.
- No retry mechanism for failed emails.

## License
This project is intended for private, internal use by authorized users of [Your Company/Website].
Unauthorized distribution or commercialization is strictly prohibited.

## Authors
- Developed and maintained by [mbasha93@github.com]
- For customization requests, please contact [mbasha@aimicromind.com]

---

