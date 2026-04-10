# M365 Mailbox Automation — Backend

Express.js API server. Deploy on Render.

## Deploy to Render

1. Push this folder to a GitHub repo
2. Go to https://render.com → New → Web Service
3. Connect your GitHub repo
4. Set these values:
   - **Build command:** `npm install`
   - **Start command:** `npm start`
   - **Environment:** Node
5. Click Deploy

## Endpoints

| Method | Path | Description |
|--------|------|-------------|
| POST | `/api/test-connection` | Validate Azure credentials |
| POST | `/api/create-mailbox` | Create a shared mailbox |
| POST | `/api/reset-password` | Reset a user's password |
| POST | `/api/add-delegation` | Add SendAs / FullAccess delegation |
| POST | `/api/enable-smtp` | Enable SMTP AUTH org-wide |
| POST | `/api/generate-script` | Download PowerShell .ps1 script |

## Required Azure App Permissions

**Microsoft Graph (Application):**
- User.ReadWrite.All
- Mail.ReadWrite
- MailboxSettings.ReadWrite
- Directory.ReadWrite.All

**Office 365 Exchange Online (Application):**
- Exchange.ManageAsApp

**Azure AD Roles (assign to app service principal):**
- Exchange Administrator
- User Administrator

After adding permissions → Grant admin consent in Azure Portal.
