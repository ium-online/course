# IUM LMS — GitHub Pages + Google Apps Script + Google Sheets + Google Drive

## Core rule
**A student can only see courses assigned to that student.** Lessons are filtered by the assigned course, and the backend verifies the student's enrollment before returning lessons, saving progress, submitting assignments, or saving quiz results.

## Architecture
- **Frontend:** GitHub Pages (`index.html` + `assets/`)
- **Backend:** Google Apps Script (`backend/Code.gs`)
- **Database:** Google Sheets
- **File storage:** Google Drive

## Roles
- **ADMIN:** approve accounts, create courses, create lessons, assign courses, view users and statistics, give feedback.
- **INSTRUCTOR:** secure login and instructor role. The instructor workspace is prepared for course/lesson management extensions.
- **STUDENT:** sees only assigned courses and their lessons, progress, quizzes, assignments and feedback.

## Setup
### 1. Google Sheet
Create a Google Sheet and copy its ID from the URL.

### 2. Google Drive
Create a folder for assignment uploads and copy the folder ID.

### 3. Apps Script
Open `backend/Code.gs` in Google Apps Script. Set:

```js
SPREADSHEET_ID: 'YOUR_SHEET_ID',
DRIVE_FOLDER_ID: 'YOUR_DRIVE_FOLDER_ID',
ADMIN_EMAIL: 'your-admin-email@example.com'
```

Run `setupSheets_()` once from Apps Script to create the required sheets.

### 4. Deploy backend
Apps Script → **Deploy → New deployment → Web app**.
- Execute as: **Me**
- Who has access: **Anyone** (or the access setting required by your institution)

Copy the `/exec` URL.

### 5. Configure GitHub frontend
Open `index.html` and replace:

```js
const API_URL='PASTE_APPS_SCRIPT_WEB_APP_EXEC_URL';
```

with your Apps Script `/exec` URL.

### 6. GitHub Pages
Upload all files to your repository with this structure:

```text
index.html
assets/IUM_LOGO.png
backend/Code.gs
README.md
```

Enable GitHub Pages from the repository's Pages settings.

## Important data model
The backend creates these sheets:

- Users
- Courses
- Lessons
- Enrollments
- Progress
- Assignments
- Submissions
- Quizzes
- QuizResults
- Feedback
- Certificates

### Course-specific access example
If student `STU001` is enrolled in `BM101` and student `STU002` is enrolled in `HR101`:

```text
STU001 → BM101 → BM101-L01, BM101-L02, BM101-L03
STU002 → HR101 → HR101-L01, HR101-L02, HR101-L03
```

STU001 cannot request HR101 lessons through the student API because the backend checks `Enrollments` before returning course content.

## Security notes
- Passwords are stored as SHA-256 hashes, not plain text.
- Login creates an expiring server-side session token.
- Student endpoints verify the authenticated student ID against the session.
- Enrollment is verified server-side for courses, lessons, assignments, quizzes, progress and feedback.
- Do not put Google Sheet IDs, Drive folder IDs, or admin secrets in public documentation beyond what is required for configuration.

## Recommended production upgrades
For a high-stakes institutional deployment, replace the simple email/password implementation with a stronger identity provider or institutional SSO, add rate limiting, audit logs, password reset, MFA, and more granular instructor permissions.
