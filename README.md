# Candidate Account Email Tracker

A Streamlit demo for managing candidate profiles in Excel and testing the
multi-account email and first-response workflow. It does **not** connect to Outlook
or send real email. Demo outbound messages and simulated replies are stored in
`demo_outlook.json`.

## Run locally

Requirements: Python 3.10 or newer.

```bash
python -m venv .venv
source .venv/Scripts/activate
python -m pip install -r requirements.txt
python -m streamlit run index.py
```

On Windows PowerShell, activate the virtual environment with
`.\.venv\Scripts\Activate.ps1`.

## Deploy on Streamlit Community Cloud

1. Push `index.py`, `requirements.txt`, and this README to a GitHub repository.
   Do not commit candidate data or generated tracker/mailbox files.
2. In Streamlit Community Cloud, create an app from that repository and select
   `index.py` as the app entry point.
3. Deploy. The app creates `candidate_tracker.xlsx` and `demo_outlook.json` in the
   app directory, using paths relative to the deployed app rather than a local
   Windows path.

The generated files are local to the running app instance. Streamlit Community
Cloud's local filesystem is not durable storage: files can be lost when an app
restarts, rebuilds, or is redeployed. Download the Excel tracker for a backup.
For persistent production records, use a database or external storage service.
Candidate data is sensitive; restrict access to the deployed app and do not commit
the generated data files to the repository.

## Demo workflow

1. Import candidate rows from an `.xlsx` or `.xlsm` file, mapping its columns, or add
   candidate profiles manually.
2. Add account names and To/CC addresses.
3. Select a candidate and multiple accounts, review the subject/body, and choose
   **Record demo emails for selected accounts**. Each account gets its own entry in the
   demo sent-mail list, with a unique `[REF:...]` subject tag. Nothing is delivered.
4. In **Track responses**, use **Simulate an account reply** to create a sample
   response. Its sender, subject, body, and received timestamp are added to the demo
   inbox in `demo_outlook.json`. You can set the received date and time to test which
   account is identified as the earliest responder.
5. Choose **Scan demo inbox file**. Replies are matched by their reference tag and
   the Excel tracker shows each account's response and the first responder per
   campaign.
6. Use **Download Excel tracker** to download the current workbook.

The Excel workbook and demo mailbox file are created beside `index.py`; both can be
downloaded from the app. The workbook must not be open in Excel while the app is
saving to it. This demo intentionally does not use Outlook, even if Outlook is
installed. Real mail sending and inbox integration require a separate supported
mail service/API for cloud deployment.
