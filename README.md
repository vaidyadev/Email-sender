📧 Email Sender App – Brief Overview
This is a Tkinter-based GUI application for sending emails with advanced features, aimed at simplifying single and bulk email communication.

✅ Key Features
Single & Bulk Email Support

Send to a single recipient manually

Or load multiple emails from Excel (auto-validates format)

Attachments

Attach multiple files (images, docs, audio, video)

Thumbnail preview of supported files in the UI

Voice Input

Compose emails using voice via speech recognition

Email Scheduling

Schedule emails for a specific date and time

Background thread monitors and sends scheduled emails

Persistent Credentials

Save and toggle visibility of email & password securely

Status Dashboard

Track total, sent, and failed emails in real-time

User-Friendly UI

Clean layout with icons, tooltips, and responsive feedback

Uses PIL for image handling and openpyxl for Excel reading

Security & Feedback

Warns about spam folder issues

Ensures SMTP authentication and graceful error handling

🧩 Tech Stack
Python

Tkinter – GUI

smtplib, EmailMessage – Sending emails

openpyxl – Reading Excel files

Pillow (PIL) – Image processing

SpeechRecognition + pygame.mixer – Voice input and audio feedback

JSON – Persistent scheduling
