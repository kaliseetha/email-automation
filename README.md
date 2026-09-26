# Candidate Account Email Tracker

A Streamlit demo for resource management teams to organize candidate profiles,
prepare separate profile messages for multiple client accounts, and identify which
account responds first. The app keeps candidate and submission data in an Excel
workbook and uses a local JSON file to demonstrate outbound messages and replies.

> **Demo only:** This version does not send email, access Outlook, or use the
> Microsoft Outlook API. The send and reply workflows are simulated and recorded in
> `demo_outlook.json`.

## Requirements for running locally

### Required software

- Windows 10 or 11.
- Python 3.10 or newer, with Python available on `PATH`.
- A modern browser such as Edge or Chrome.
- This project folder, including `index.py` and `requirements.txt`.

### Optional software

- Git, if you want to clone the project from GitHub.
- Microsoft Excel, if you want to open or inspect the generated workbook. Excel is
  not required to run the app. Close the tracker workbook in Excel before the app
  saves changes.
- Outlook is **not required** for this demo. Installing Outlook does not enable
  sending or reading messages in this version.

## Install and run on Windows

Open Git Bash or PowerShell in the project folder. If using Git Bash, use:

```bash
cd /e/Learning/Streamlit/email-automation
python --version
```

If `python` is not recognized, install Python from
[python.org/downloads/windows](https://www.python.org/downloads/windows/) and enable
**Add python.exe to PATH**, or add the Python installation and its `Scripts` folder
to your Windows user `PATH`. Close and reopen the terminal, then verify
`python --version` works.

Create a virtual environment, activate it, install the app dependencies, and start
Streamlit:

```bash
python -m venv .venv
source .venv/Scripts/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run index.py
```

PowerShell activation instead:

```powershell
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation, allow it for the current terminal session and
activate again:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

Open the **Local URL** printed by Streamlit in your browser. To stop the app, press
`Ctrl+C` in the terminal.

## Where the demo data is stored

The app creates these files beside `index.py`:

- `candidate_tracker.xlsx` — candidate, account, sent-demo, and response tracking
  records.
- `demo_outlook.json` — simulated outbound messages and simulated inbox replies.

The app provides download buttons for both files. These files are ignored by Git so
candidate data is not accidentally committed. Keep backups in an approved, access
controlled location; candidate profiles can contain personal information.

## Outlook: what works and what requires approval

### Current version

The current version has **no Outlook integration**. It records a pretend sent
message in `demo_outlook.json`; the response screen lets you create simulated
replies and scan the demo inbox file. No message leaves the app and no real Outlook
mailbox is read.

### If Outlook is installed on a local computer

Outlook being installed does not, by itself, connect this app to the mailbox. Real
mail sending and reading need an explicitly implemented integration, authentication,
and organizational approval. This demo does not contain that integration.

For an organization using Microsoft 365, the recommended direction is normally
the Microsoft Graph mail API (the Microsoft-supported API for Outlook/Exchange
mail), with Microsoft Entra ID sign-in. The Outlook desktop client is not required
for Graph-based mail access, and installing it does not grant API access.

Before implementation, obtain approval from the resource-management/business owner,
security/privacy team, and Microsoft 365/Entra administrator. The expected software
and service prerequisites are:

- A Microsoft 365 work account and an Exchange Online mailbox for each user or an
  explicitly approved shared-mailbox design.
- A Microsoft Entra ID app registration and a supported OAuth sign-in flow, such as
  Microsoft Authentication Library (MSAL).
- A registered application/redirect URI and securely managed client configuration.
  Do not put client secrets or credentials in source code or a public repository.
- Admin review and consent for the least-privilege Microsoft Graph mail permissions.
  A delegated user-sign-in design commonly needs `Mail.Send` to send and `Mail.Read`
  to scan replies; exact permissions and mailbox scope must be confirmed by the
  Microsoft 365 administrator for the tenant and chosen design.
- An approved hosting location for a cloud deployment, with TLS, restricted access,
  secure authentication, and an approved data-retention and audit policy.

Sending and reading mail can involve sensitive candidate and mailbox data. Agree on
mailbox scope, retention, audit, and access controls before enabling it. Tenant
consent and approval requirements vary; this demo does not request or use any of
these permissions.

An alternative local-only approach is Windows COM automation through **classic
Outlook**. That approach requires classic Outlook installed and configured in the
same signed-in Windows session as the app, plus the Windows COM integration
dependency. It is not the Outlook web/API integration, is not suitable for Streamlit
Community Cloud, and is not implemented in this demo. The new Outlook client may not
support that COM automation.

### Can the cloud app use Outlook installed on my PC?

No. A Streamlit Community Cloud app runs on a remote server; it cannot access Outlook
installed on your office computer. Installing Outlook locally does not give the
cloud server mailbox access. A cloud deployment needs a supported server-side mail
integration, such as Microsoft Graph, with the required tenant approval and
authentication configured securely.

## Sharing with multiple team members

### Streamlit Community Cloud

1. Push the source files (`index.py`, `requirements.txt`, and `README.md`) to a
   GitHub repository. Do not commit the generated tracker or demo mailbox.
2. Create a Streamlit Community Cloud app from that repository and select `index.py`
   as the entry point.
3. Restrict app and repository access to the intended users. The demo currently has
   no user authentication or role-based access controls.

Community Cloud's local filesystem should not be treated as permanent storage.
Generated files can be lost on restart, rebuild, or redeploy, and local Excel/JSON
files are not a safe shared multi-user database. Download backups for a demo. For a
team production deployment, plan for centrally managed storage (for example, a
database), authentication/authorization, backups, and a supported mail integration.

### Can this be distributed as an EXE?

It is possible to package a Python/Streamlit app for Windows using a bundling tool
such as PyInstaller, but this app has not been packaged or tested as an EXE. Streamlit
still needs to run a local web server and open a browser; packaging dependencies,
browser launch, Windows security controls, and updates require additional testing.
Each user's EXE would normally keep its own local workbook and JSON file, so users
would not automatically share data. An EXE alone does not provide a central,
concurrent team workspace or connect Outlook.

For multiple team members working on the same candidates and responses, a centrally
hosted app with authentication and shared persistent storage is generally a better
fit than distributing separate EXEs.

## Demo workflow

1. Add candidates manually or import candidate rows from an `.xlsx` or `.xlsm` file.
2. Add account names and To/CC email addresses.
3. Select a candidate and multiple accounts, review the subject and body, and choose
   **Record demo emails for selected accounts**. The app records a separate
   reference-tagged message for each account in `demo_outlook.json`; it does not send
   them.
4. In **Track responses**, use **Simulate an account reply** and set a responder and
   received time. The reply is saved to the demo inbox file.
5. Choose **Scan demo inbox file**. The app matches replies by their unique subject
   reference and identifies the earliest account response in the campaign.
6. Download the Excel tracker and demo mailbox JSON if you need a copy of the demo
   data.
