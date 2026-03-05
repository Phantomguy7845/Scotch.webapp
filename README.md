# Scotch webapps1

Two-page web app for company vehicle requests:

- `index.html`: Landing page + request form
- `admin.html`: Admin approval page
- Backend API + database: Google Apps Script + Google Sheets
- Hosting: GitHub Pages (frontend)

## Project Structure

```text
Scotch_webapps1/
  index.html
  admin.html
  assets/
    css/styles.css
    js/config.js
    js/app.js
    js/admin.js
  apps-script/
    Code.gs
```

## 1) Setup Google Sheets + Apps Script

1. Create a new Google Sheet.
2. Open `Extensions > Apps Script`.
3. Replace default script content with [`apps-script/Code.gs`](apps-script/Code.gs).
4. Update this line in Apps Script:
   - `ADMIN_KEY: "CHANGE_ME_TO_A_STRONG_ADMIN_KEY"`
5. Save the script project.
6. Deploy as web app:
   - `Deploy > New deployment > Web app`
   - Execute as: `Me`
   - Who has access: `Anyone`
7. Copy the deployed web app URL (ends with `/exec`).

## 2) Configure Frontend

1. Open [`assets/js/config.js`](assets/js/config.js).
2. Update:
   - `apiBaseUrl`: paste your Apps Script `/exec` URL
   - `adminKey`: optional, or leave empty and type it on admin page
3. Save file.

## 3) Deploy Frontend to GitHub Pages

1. Create a new GitHub repository and push this project.
2. In GitHub repository settings:
   - `Pages > Build and deployment`
   - Source: `Deploy from a branch`
   - Branch: `main` (root)
3. Wait for page publishing.

Your URLs will look like:

- `https://<github-username>.github.io/<repo-name>/index.html`
- `https://<github-username>.github.io/<repo-name>/admin.html`

## 4) Test Flow

1. Open `index.html` and submit a request.
2. Open `admin.html`, enter admin key, and refresh requests.
3. Click `Approve and notify`.
4. Verify:
   - Status changes to `APPROVED` in admin page.
   - Row in Google Sheet updates with `approvedAt`, `approvedBy`, `emailStatus`.
   - Requester receives approval email.

## Important Notes

- This is a lightweight admin-key model. For stronger security, add proper authentication (Google Sign-In, OAuth, or server-side auth).
- Do not expose sensitive secrets in public repositories.
- Apps Script quotas apply for email sending (`MailApp`).
