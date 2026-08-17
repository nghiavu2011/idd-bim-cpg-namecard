# IDD BIM CPG Namecard — PoC Setup v1.0

## Fixed resources
- Google Sheet ID: `13XsLUK792UkdDuXYHSdjIGJPX7AazQt6hROtTa0LRCI`
- Master tab: `Namecard_CRM_Contacts`
- Google Drive folder ID: `1Sy6Lb509_seEfP-EQd_1Wh_1lipWIxG3`

## 1. Install the Google Apps Script backend
1. Open the Google Sheet.
2. Choose **Extensions → Apps Script**.
3. Replace `Code.gs` with `IDD_BIM_CPG_Namecard_AppsScript_v1.0.gs`.
4. Save.
5. Run `setupSystem()` once from the Apps Script editor.
6. Approve Google permissions when prompted.
7. Open **Execution log** and copy the generated `path_secret`.

`setupSystem()` will create these tabs if they do not exist:
- `Cards_Inbox`
- `Review_Queue`

It will not reorder or overwrite the existing 27-column contact master.

## 2. Deploy as Web App
1. Apps Script → **Deploy → New deployment**.
2. Type: **Web app**.
3. Execute as: **Me**.
4. Who has access: **Anyone** for this PoC.
5. Deploy and copy the `/exec` URL.

The Action server URL is:

`https://script.google.com/macros/s/AKfycbxO-vRSFGHc8IU6esuACbXmabInZwpem2yloquIsawcMuoBg5cRklKsMow86JHpNuai/exec/nmc_2a69e51629fb43639410e863`

The path secret is an obscurity layer for the PoC, not enterprise-grade authentication.

## 3. Configure the Custom GPT
1. Create/edit **IDD BIM CPG Namecard** in the GPT Builder.
2. Paste `IDD_BIM_CPG_Namecard_MasterPrompt_v1.0.md` into Instructions.
3. Use **Actions**, not Apps. Apps and Actions are mutually exclusive in GPT configuration.
4. Create a new Action.
5. Authentication: **None** for this PoC.
6. Paste `IDD_BIM_CPG_Namecard_Action_v1.0.yaml`.
7. Replace `REPLACE_DEPLOYMENT_ID` and `REPLACE_PATH_SECRET` in the server URL.
8. Test the Action in Preview.
9. Share the GPT internally with chat/use permission only.

## 4. PoC test cases
Use three cards:

### Test A — New contact
Upload a card for a person not in the 17-contact master.
Expected:
- `NEW`
- New `C###` row in `Namecard_CRM_Contacts`
- New `NC#####` row in `Cards_Inbox`

### Test B — Exact duplicate
Upload an existing person's card with the same primary email or mobile.
Expected:
- `MERGED`
- No new master contact
- `Số card trùng` increases by 1
- New `NC#####` row in `Cards_Inbox`

### Test C — Ambiguous match
Upload a card whose name/company resembles an existing contact but lacks exact phone/email match.
Expected:
- `REVIEW`
- New row in `Cards_Inbox`
- New `RV#####` row in `Review_Queue`
- No automatic merge

## 5. Important image-archive test
The backend already supports saving an image when the Action actually receives either:
- `image_base64` + `image_mime_type`, or
- a fetchable `image_url`.

The current Custom GPT Action documentation does not clearly guarantee that the raw bytes of a user-uploaded chat image are automatically forwarded to an Action. Therefore the PoC must verify this separately.

Do **not** let the GPT invent a URL/base64 value. If image bytes are unavailable, the data workflow still completes and `Cards_Inbox` records `IMAGE_NOT_RECEIVED_BY_ACTION` via the returned status. A later phase can solve image archiving through a dedicated upload tool or agent-owned Google Drive connection if required.

## 6. After the PoC passes
Next production steps:
- stop relying on public-edit Google resources;
- put a thin authenticated proxy in Vercel/Cloudflare if stronger Action authentication is needed;
- connect the CRM HTML to a read-only JSON endpoint;
- add passcode/session protection;
- poll `data_version` every 15–30 seconds and refresh only when the version changes;
- scale UI typography so browser 100% visually matches the current preferred 125% view.
