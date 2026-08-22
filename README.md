# 📅 Automated Life Coach Appointment Scheduler & Sync Engine

An enterprise-grade, lightweight scheduling and synchronization engine built on **Google Apps Script**, **Google Forms**, **Google Sheets**, and **Gmail API**.

This system automatically ingests availability schedules sent via email by a Life Coach, updates a live Google Form choice list in real time, prevents race-condition double bookings using script locks, generates Google Calendar invites with Google Meet links, and automates employee email broadcasts.

---

## 🏗️ Architecture & Workflow

[Coach Email] ──> (checkAndParseCoachEmails) ──> [Master Log Sheet] <── (onEmployeeFormSubmit) <── [Employee Form]
                                                     │
                                                     ├──> [Google Form Sync Engine]
                                                     ├──> [Google Calendar Invite API]
                                                     └──> [Mon/Wed Broadcast Filter]

1. **Email Parsing Engine (`checkAndParseCoachEmails`)**: Periodically checks Gmail for unread schedule updates sent from the Coach's verified address. Parses dynamic Latvian day names (e.g., *Pirmdiena*, *Otrdiena*), explicit date overrides (e.g., `24.08.`), and preserves 30-minute start time offsets (e.g., `18:30 - 19:30`).
2. **Reconciliation Engine (`reconcileFollowUpSlots`)**: Compares incoming email schedules against historical records in the `Master Log` sheet. New slots are added, removed slots are marked as `Cancelled by Coach`, and active employee bookings are strictly protected.
3. **Concurrency Protection (`onEmployeeFormSubmit`)**: Form submissions execute within a `LockService` thread block to prevent double-booking race conditions when multiple users submit simultaneously.
4. **Google Calendar API Integration**: Generates a 1-hour Google Calendar event complete with a Google Meet video conference link, sending automatic calendar invites to both the Coach and the Employee.
5. **Selective Broadcast Filter**: Triggers HTML broadcast emails to employees on Mondays and Wednesdays while silently updating form options on Tuesdays, Thursdays, Fridays, and weekends to avoid email fatigue.

---

## 📊 Database Schema (`Master Log` Sheet)

The Google Sheet tab named **`Master Log`** must contain the following column structure:

| Column | Header Name | Description | Example Data |
| :--- | :--- | :--- | :--- |
| **A** | `Unique ID` | Normalized date and start time lookup key | `'17.08.2026. 15:00` |
| **B** | `Date` | Formatted calendar date | `17.08.2026.` |
| **C** | `Time Window` | 1-Hour calculated window | `15:00 - 16:00` |
| **D** | `Status` | Slot lifecycle state | `Available` / `Booked` / `Cancelled by Coach` / `Expired` |
| **E** | `Employee Name` | Name of attendee | `Jane Doe` |
| **F** | `Employee Email` | Email address of attendee | `jane.doe@example.com` |
| **G** | `Event ID` | Google Calendar API Event ID | `_60o30c1g60o30c1g...` |

---

## ⚙️ Environment Configuration

Set the three configuration constants at the top of `Code.gs`:

    const EMPLOYEE_FORM_ID = 'YOUR_EMPLOYEE_FORM_ID_HERE'; // Live Employee Google Form ID
    const COACH_EMAIL      = 'coach@example.com';          // Coach's verified email address
    const HR_EMAIL         = 'hr-notifications@example.com'; // Broadcast distribution list

---

## ⏰ Google Apps Script Trigger Setup

To deploy the automation, open the Google Apps Script project attached to your Master Log spreadsheet and configure the following triggers (**Triggers ➔ Add Trigger**):

| Function | Event Source | Event Type | Details |
| :--- | :--- | :--- | :--- |
| `onEmployeeFormSubmit` | From spreadsheet | On form submit | Triggers when an employee submits the booking form |
| `checkAndParseCoachEmails` | Time-driven | Hour timer | Runs every hour (or every 30 mins) to process emails |

---

## 🛡️ Robustness & Safety Features

* **Double-Booking Prevention**: Uses `LockService.getScriptLock()` with a 10-second wait lock. If a user loses a booking race condition, they receive an automated HTML apology email asking them to re-submit for another time.
* **Auto-Expiration (`archivePastSlots`)**: Past unbooked slots (`Date < Today`) are automatically updated to `Expired` prior to form updates, preventing past dates from lingering in live dropdown choices.
* **Deduplication Guard**: Form choice arrays are passed through `Array.from(new Set(...))` to guarantee Google Forms never throws duplicate option errors.
* **Date Parsing Bugfix**: Explicit date extraction (`\b(\d{1,2})\.(\d{1,2})\.`) is strictly constrained to text *before* the dash separator, ensuring time tokens (e.g., `10.00`) are never misparsed as dates.

---

## 📋 Operational Manual: Handling Cancellations

If an employee requests to cancel a booking manually:
1. Open **Google Calendar** and delete the meeting event.
2. Open the **`Master Log`** tab in your Google Sheet.
3. Locate the slot row and change **Column D (Status)** from `Booked` back to **`Available`**.
4. Clear Columns **E**, **F**, and **G** (Name, Email, Event ID).
5. Run `updateFormChoices()` manually from Apps Script (or wait for the next automated execution to republish the slot live).
