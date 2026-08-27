ATHARV EXAM MANAGEMENT SYSTEM V8 - AUTO RESULT NOTIFICATIONS

NEW:
1. Student submits exam.
2. Result, percentage and rank are calculated.
3. Result email is automatically sent if Gmail settings are configured.
4. WhatsApp notification is sent automatically when an approved WhatsApp API is configured.

SETUP:
Admin -> Email Settings
- Gmail address
- 16-character Gmail App Password

Admin -> WhatsApp & Auto Result
- Keep Auto Result Email enabled
- Optional: enable WhatsApp and enter your approved provider/API endpoint and API token.

IMPORTANT:
WhatsApp requires an official Business/API provider configuration. The exact API endpoint/token depend on your provider.

Run RUN_EXAM_SYSTEM.bat after extracting.


V10 LICENSE MODULE ADDED:
- Unique license generation
- Institute/customer details
- Device binding
- Pending/Active/Blocked status
- Expiry and lifetime support
- Renewal and device reset from License Admin

V10 USER LOGIN MODULE: License-wise Institute Admin username/password creation.

AUTO SCHOOL LOGIN: Every new license automatically creates a unique Institute Admin User ID and Password.

EMAIL LOGIN DELIVERY
1. Login as Super Admin.
2. Open Email Settings.
3. Enter SMTP host, port, sender email and App Password.
4. Enable Automatic Email.
5. While generating a license, enter the School/College email address.
6. User ID, Password and License details will be emailed automatically.

FINAL V2: app.py syntax verified. License email field indentation fixed.

LASTROWID FIX: License INSERT now stores the execute cursor and reads cursor.lastrowid. Syntax verified.

INSTITUTE ADMIN MODULE
- License-linked Institute Admin account
- Change own password
- Create Teacher/Staff users
- Reset Teacher/Staff passwords
- Delete Teacher/Staff users
- Institute backup download
- School/College name shown in header
- Footer branding: Developed by Arjun Kakde | 7758053091 / 7775804777
