/**
 * Automated Life Coach Appointment Scheduler & Sync Engine
 * 
 * Features:
 * - Gmail ingestion engine that parses schedule updates directly from the Coach's incoming emails.
 * - Bidirectional reconciliation (adds new slots, cancels removed unbooked slots, protects booked slots).
 * - Regex-based parser handling Latvian day names, explicit dates, and 30-minute time offsets.
 * - Real-time Google Form synchronization with LockService concurrency protection.
 * - Selective broadcast filter (silent updates on Tue/Thu/Fri; broadcast emails on Mon/Wed).
 */

// ==========================================
// CONFIGURATION CONSTANTS
// ==========================================
const EMPLOYEE_FORM_ID = 'YOUR_EMPLOYEE_FORM_ID_HERE'; // Replace with live Employee Google Form ID
const COACH_EMAIL      = 'coach@example.com';          // Replace with Coach's email address
const HR_EMAIL         = 'hr-notifications@example.com'; // Replace with HR / Broadcast distribution email

// ==========================================
// 1. TRIGGER: EMPLOYEE BOOKING SUBMISSION
// ==========================================
function onEmployeeFormSubmit(e) {
  const lock = LockService.getScriptLock();
  
  try {
    // Acquire script lock for up to 10 seconds to eliminate double-booking race conditions
    lock.waitLock(10000); 
    
    const responses     = e.values; 
    const timestamp     = responses[0];
    const employeeName  = responses[2]; 
    const employeeEmail = responses[3]; 
    const chosenSlot    = responses[4]; 
    
    Logger.log(`Processing booking request for ${employeeName} at slot: ${chosenSlot}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterLogSheet = ss.getSheetByName('Master Log');
    const data = masterLogSheet.getDataRange().getDisplayValues(); 
    
    let slotFound = false;
    let rowIndex = -1;
    
    // Scan Master Log database for slot availability
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].trim() === chosenSlot.trim() && data[i][3] === 'Available') {
        slotFound = true;
        rowIndex = i + 1;
        break;
      }
    }
    
    // Guard Clause: Handle rare race condition where slot was taken while user filled the form
    if (!slotFound) {
      const apologyHtml = `
        <div style="font-family: Arial, sans-serif; padding: 20px; color: #333;">
          <h3 style="color: #d32f2f;">Atvainojiet! / Sorry!</h3>
          <p>Diemžēl izvēlētais laiks (<strong>${chosenSlot}</strong>) tikko tika rezervēts.<br>
          Lūdzu, atgriezieties anketā un izvēlieties citu pieejamo laiku.</p>
          <hr style="margin: 15px 0;">
          <p>Unfortunately, the time slot you selected (<strong>${chosenSlot}</strong>) was just booked by someone else.<br>
          Please return to the form and select another available time.</p>
        </div>
      `;
      GmailApp.sendEmail(employeeEmail, "Life Coach Rezervācija / Booking - Kļūda / Error", "", { htmlBody: apologyHtml });
      return;
    }
    
    // Create Calendar Event & Update Database Record
    const eventId = createCalendarEvent(chosenSlot, employeeName, employeeEmail);
    
    masterLogSheet.getRange(rowIndex, 4).setValue('Booked');
    masterLogSheet.getRange(rowIndex, 5).setValue(employeeName);
    masterLogSheet.getRange(rowIndex, 6).setValue(employeeEmail);
    masterLogSheet.getRange(rowIndex, 7).setValue(eventId);
    
    // Refresh Google Form to instantly remove the booked slot
    updateFormChoices();
    
  } catch (error) {
    Logger.log('CRITICAL ERROR inside onEmployeeFormSubmit: ' + error.toString());
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 2. ENGINE: GMAIL INGESTION & PARSER
// ==========================================
function checkAndParseCoachEmails() {
  Logger.log("Starting Gmail check for Life Coach schedule updates...");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterLogSheet = ss.getSheetByName('Master Log');
  
  // 1. Auto-expire past unbooked slots prior to sync
  archivePastSlots(masterLogSheet);
  
  // 2. Search for unread schedule emails sent strictly by the Coach
  const threads = GmailApp.search(`label:unread from:${COACH_EMAIL} subject:("Šīs nedēļas logi" OR "Laiku apdeiti" OR "apdeit")`);
  
  if (threads.length === 0) {
    Logger.log("No new unread schedule emails found.");
    return;
  }
  
  // Process latest message thread
  const thread = threads[0];
  const message = thread.getMessages()[thread.getMessageCount() - 1]; 
  const bodyText = message.getPlainBody();
  const subject = message.getSubject();
  const emailDate = message.getDate();
  
  Logger.log(`Found email with subject: "${subject}" received on ${emailDate}`);
  
  // Parse slot text payload
  const slotsFromEmail = extractSlotsFromText(bodyText, emailDate);
  Logger.log(`Successfully parsed ${slotsFromEmail.length} slots from email body.`);
  
  // 3. Reconcile changes (adds new slots, cancels unbooked missing slots, protects booked slots)
  reconcileFollowUpSlots(slotsFromEmail, masterLogSheet);
  updateFormChoices();
  
  // 4. Selective Broadcast Filter: Only send emails on Mondays (1) and Wednesdays (3)
  const dayOfWeek = emailDate.getDay(); // 0 = Sun, 1 = Mon, 2 = Tue, 3 = Wed, 4 = Thu, 5 = Fri, 6 = Sat
  const isExplicitUpdate = subject.includes('2') || 
                           subject.toLowerCase().includes('apdeit') || 
                           subject.toLowerCase().includes('izmaiņ');

  if (dayOfWeek === 1 && !isExplicitUpdate) {
    Logger.log("Monday email detected: Triggering Monday Broadcast...");
    sendMondayBroadcast();
  } else if (dayOfWeek === 3 || (dayOfWeek === 1 && isExplicitUpdate)) {
    Logger.log("Wednesday/Update email detected: Triggering Wednesday Broadcast...");
    sendWednesdayReminder();
  } else {
    Logger.log(`Update processed silently on day ${dayOfWeek} (Tue/Thu/Fri/Weekend). Form choices synced without broadcast.`);
  }
  
  // Mark email thread as read
  thread.markRead();
  Logger.log("Email processed and thread marked as read.");
}

// ==========================================
// HELPER: CREATE GOOGLE CALENDAR INVITE
// ==========================================
function createCalendarEvent(chosenSlot, employeeName, employeeEmail) {
  const parts = chosenSlot.trim().split(' ');
  const dateString = parts[0].replace(/\.+$/, '');
  const dateParts = dateString.split('.');
  const timeParts = parts[1].split(':');
  
  const day     = parseInt(dateParts[0], 10);
  const month   = parseInt(dateParts[1], 10) - 1; // JS Months 0-11
  const year    = parseInt(dateParts[2], 10);
  const hours   = parseInt(timeParts[0], 10);
  const minutes = parseInt(timeParts[1], 10);
  
  const startDateTime = new Date(year, month, day, hours, minutes);
  const endDateTime = new Date(startDateTime.getTime() + (60 * 60 * 1000)); // 1-hour session duration
  
  const options = {
    description: `Life Coach session for employee ${employeeName}.`,
    guests: `${COACH_EMAIL},${employeeEmail}`,
    sendInvites: true
  };
  
  const event = CalendarApp.getDefaultCalendar().createEvent(
    `Life Coach Session - ${employeeName}`,
    startDateTime,
    endDateTime,
    options
  );
  
  return event.getId(); 
}

// ==========================================
// HELPER: UPDATE EMPLOYEE FORM CHOICES
// ==========================================
function updateFormChoices() {
  const form = FormApp.openById(EMPLOYEE_FORM_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterLogSheet = ss.getSheetByName('Master Log');
  const data = masterLogSheet.getDataRange().getDisplayValues();
  
  let availableSlots = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'Available') { 
      availableSlots.push(data[i][0]); 
    }
  }
  
  // Enforce choice uniqueness before updating Google Form
  availableSlots = Array.from(new Set(availableSlots));
  
  const items = form.getItems();
  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle().includes('Available time slots') || items[i].getTitle().includes('Pieejamie laiki')) {
      const multipleChoiceQuestion = items[i].asMultipleChoiceItem(); 
      if (availableSlots.length > 0) {
        multipleChoiceQuestion.setChoiceValues(availableSlots);
      } else {
        multipleChoiceQuestion.setChoiceValues(['No sessions available right now / Nav pieejamu sesiju']);
      }
      break;
    }
  }
}

// ==========================================
// HELPER: AUTOMATICALLY EXPIRE PAST SLOTS
// ==========================================
function archivePastSlots(sheet) {
  const data = sheet.getDataRange().getDisplayValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'Available') {
      let dateStr = data[i][1].replace(/\.+$/, '');
      let parts = dateStr.split('.');
      if (parts.length === 3) {
        let day = parseInt(parts[0], 10);
        let month = parseInt(parts[1], 10) - 1;
        let year = parseInt(parts[2], 10);
        let slotDate = new Date(year, month, day);
        if (slotDate < today) {
          sheet.getRange(i + 1, 4).setValue('Expired');
        }
      }
    }
  }
}

// ==========================================
// HELPER: TEXT PARSING & REGEX ENGINE
// ==========================================
function extractSlotsFromText(text, referenceDate) {
  const dayOfWeek = referenceDate.getDay(); 
  const distanceToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
  const mondayOfWeek = new Date(referenceDate.getTime());
  mondayOfWeek.setDate(referenceDate.getDate() + distanceToMonday);
  
  const dayMap = {
    'pirmdiena': 0, 'otrdiena': 1, 'trešdiena': 2, 'ceturtdiena': 3, 'piektdiena': 4
  };
  
  const lines = text.split('\n');
  let currentSection = "this_week"; 
  let parsedSlots = [];
  
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].toLowerCase().trim();
    if (line.includes('nākamā nedēļa') || line.includes('nakama nedela')) {
      currentSection = "next_week";
      continue;
    }
    if (line.includes('šīs nedēļas logi') || line.includes('sis nedelas logi')) {
      currentSection = "this_week";
      continue;
    }
    
    let foundDay = null;
    for (let day in dayMap) {
      if (line.includes(day)) {
        foundDay = day;
        break;
      }
    }
    
    if (foundDay) {
      let standardLine = lines[i].replace('–', '-');
      if (!standardLine.includes('-')) continue;
      
      let lineParts = standardLine.split('-');
      let dayPart = lineParts[0];
      let timePart = lineParts.slice(1).join('-');
      
      let timeMatches = timePart.match(/\b\d{1,2}[.:]\d{2}\b/g);
      if (!timeMatches) continue;
      
      let targetDate;
      // Search for explicit date (e.g., "(24.08.)") strictly BEFORE the dash
      let explicitDateMatch = dayPart.match(/\b(\d{1,2})\.(\d{1,2})\.?(?:(\d{4}))?\b/);
      
      if (explicitDateMatch) {
        let dayNum   = parseInt(explicitDateMatch[1], 10);
        let monthNum = parseInt(explicitDateMatch[2], 10) - 1; 
        let yearNum  = explicitDateMatch[3] ? parseInt(explicitDateMatch[3], 10) : referenceDate.getFullYear();
        targetDate   = new Date(yearNum, monthNum, dayNum);
      } else {
        let daysToAdd = dayMap[foundDay];
        if (line.includes('nākamā') || line.includes('nakama') || currentSection === "next_week") {
          daysToAdd += 7;
        }
        targetDate = new Date(mondayOfWeek.getTime());
        targetDate.setDate(mondayOfWeek.getDate() + daysToAdd);
      }
      
      let dd = String(targetDate.getDate()).padStart(2, '0');
      let mm = String(targetDate.getMonth() + 1).padStart(2, '0');
      let yyyy = targetDate.getFullYear();
      let formattedDate = `${dd}.${mm}.${yyyy}.`;
      
      for (let j = 0; j < timeMatches.length; j++) {
        let cleanTime = timeMatches[j].replace('.', ':');
        let timeParts = cleanTime.split(':');
        let startHour = parseInt(timeParts[0], 10);
        let startMins = timeParts[1]; 
        let endTime = `${startHour + 1}:${startMins}`; // Preserves minute offsets (e.g., 18:30 -> 19:30)
        
        parsedSlots.push({
          uniqueId: `'${formattedDate} ${cleanTime}`,
          date: formattedDate,
          timeWindow: `${cleanTime} - ${endTime}`
        });
      }
    }
  }
  return parsedSlots;
}

// ==========================================
// RECONCILIATION ENGINE
// ==========================================
function reconcileFollowUpSlots(emailSlots, sheet) {
  const existingData = sheet.getDataRange().getDisplayValues();
  
  // Append newly added slots
  emailSlots.forEach(eSlot => {
    let found = false;
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] === eSlot.uniqueId.replace("'", "")) {
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([eSlot.uniqueId, eSlot.date, eSlot.timeWindow, 'Available', '', '', '']);
    }
  });
  
  // Cancel unbooked slots missing from the updated email list
  for (let i = 1; i < existingData.length; i++) {
    let sheetSlotId = existingData[i][0];
    let sheetStatus = existingData[i][3];
    
    if (sheetStatus === 'Available') {
      let stillExistsInEmail = emailSlots.some(eSlot => eSlot.uniqueId.replace("'", "") === sheetSlotId);
      if (!stillExistsInEmail) {
        sheet.getRange(i + 1, 4).setValue('Cancelled by Coach');
      }
    }
  }
}

// ==========================================
// BROADCAST HANDLER: MONDAYS
// ==========================================
function sendMondayBroadcast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterLogSheet = ss.getSheetByName('Master Log');
  archivePastSlots(masterLogSheet);
  
  const data = masterLogSheet.getDataRange().getDisplayValues();
  const form = FormApp.openById(EMPLOYEE_FORM_ID);
  const formUrl = form.getPublishedUrl();
  
  let availableSlots = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'Available') {
      availableSlots.push(`${data[i][1]} | ${data[i][2]}`);
    }
  }
  
  if (availableSlots.length === 0) return;
  
  const subject = "Life coach - šonedēļ / this week";
  let slotsHtml = "<ul style='font-family: Arial, sans-serif; color: #333; line-height: 1.6; padding-left: 20px;'>";
  availableSlots.forEach(slot => {
    slotsHtml += `<li style='margin-bottom: 6px;'>&#128197; <strong>${slot}</strong></li>`;
  });
  slotsHtml += "</ul>";
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 650px; border: 1px solid #e0e0e0; padding: 25px; border-radius: 8px; color: #333; line-height: 1.5;">
      <p>Sveiki,</p>
      <p>Sūtu jums informāciju par šīs nedēļas Life coach sesiju laikiem.<br>
      Sesijas norisinās attālināti, izmantojot <strong>Google Meet</strong>.</p>
      <p style="font-size: 13px; color: #555;">
        *Ja darbinieks strādā no biroja un ir nepieciešama telpa priekš life coach sesijas, lūgums to norādīt e-pastā.<br>
        *Tiem, kuri ir mājās, jāpārliecinās, vai viņiem ir iespēja redzēt un dzirdēt life coach attālinātajā sesijā caur Google Meet.
      </p>
      <p style="margin-top: 20px;"><strong>Pieejamie sesiju laiki tuvākajām dienām:</strong></p>
      ${slotsHtml}
      <div style="margin: 25px 0; text-align: center;">
        <a href="${formUrl}" style="background-color: #673ab7; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">Pieteikties sesijai šeit / Book session here</a>
      </div>
      <hr style="border: 0; border-top: 1px solid #eeeeee; margin: 30px 0;">
      <p>Hello!</p>
      <p>Here you will find the information about this week's Life coach session times.<br>
      The sessions take place remotely using <strong>Google Meet</strong>.</p>
      <p style="font-size: 13px; color: #555;">
        *If the employee works from the office and needs a room for a life coach session, please indicate this in email.<br>
        *Those who are home should make sure they are able to see and hear the life coach in a remote session via Google Meet.
      </p>
      <p style="margin-top: 20px;"><strong>These are the available session times for this week:</strong></p>
      ${slotsHtml}
      <div style="margin: 25px 0; text-align: center;">
        <a href="${formUrl}" style="background-color: #673ab7; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">Sign up here / Book session here</a>
      </div>
      <p style="margin-top: 30px; font-size: 11px; color: #999; text-align: center;">Šis ir automātisks paziņojums. / This is an automated notification.</p>
    </div>
  `;
  
  GmailApp.sendEmail(HR_EMAIL, subject, "", { htmlBody: htmlBody });
}

// ==========================================
// BROADCAST HANDLER: WEDNESDAYS
// ==========================================
function sendWednesdayReminder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterLogSheet = ss.getSheetByName('Master Log');
  archivePastSlots(masterLogSheet);
  
  const data = masterLogSheet.getDataRange().getDisplayValues();
  const form = FormApp.openById(EMPLOYEE_FORM_ID);
  const formUrl = form.getPublishedUrl();
  
  let availableSlots = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'Available') {
      availableSlots.push(`${data[i][1]} | ${data[i][2]}`);
    }
  }
  
  if (availableSlots.length === 0) return;
  
  const subject = "Atjaunots! Life coach sesijas - šonedēļ! / Updated! Sessions - this week!";
  let slotsHtml = "<ul style='font-family: Arial, sans-serif; color: #333; line-height: 1.6; padding-left: 20px;'>";
  availableSlots.forEach(slot => {
    slotsHtml += `<li style='margin-bottom: 6px;'>&#128197; <strong>${slot}</strong></li>`;
  });
  slotsHtml += "</ul>";
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 650px; border: 1px solid #e0e0e0; padding: 25px; border-radius: 8px; color: #333; line-height: 1.5;">
      <p>Sveiki,</p>
      <p>Sūtu jums atjaunoto informāciju par šīs nedēļas Life coach sesiju laikiem.<br>
      Sesijas norisinās attālināti, izmantojot <strong>Google Meet</strong>.</p>
      <p style="font-size: 13px; color: #555;">
        *Ja darbinieks strādā no biroja un ir nepieciešama telpa priekš Life coach sesijas, lūgums to norādīt e-pastā.<br>
        *Tiem, kuri ir mājās, jāpārliecinās, vai viņiem ir iespēja redzēt un dzirdēt life coach attālinātajā sesijā caur Google Meet.
      </p>
      <p style="margin-top: 20px;"><strong>Pieejamie sesiju laiki tuvākajām dienām:</strong></p>
      ${slotsHtml}
      <div style="margin: 25px 0; text-align: center;">
        <a href="${formUrl}" style="background-color: #673ab7; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">Pieteikties sesijai šeit / Book session here</a>
      </div>
      <hr style="border: 0; border-top: 1px solid #eeeeee; margin: 30px 0;">
      <p>Hello!</p>
      <p>Here you will find updated information about this week's Life coach session times.<br>
      The sessions take place remotely using <strong>Google Meet</strong>.</p>
      <p style="font-size: 13px; color: #555;">
        *If the employee works from the office and needs a room for a life coach session, please indicate this in email.<br>
        *Those who are home should make sure they are able to see and hear the life coach in a remote session via Google Meet.
      </p>
      <p style="margin-top: 20px;"><strong>These are the available session times for this week:</strong></p>
      ${slotsHtml}
      <div style="margin: 25px 0; text-align: center;">
        <a href="${formUrl}" style="background-color: #673ab7; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">Sign up here / Book session here</a>
      </div>
      <p style="margin-top: 30px; font-size: 11px; color: #999; text-align: center;">Šis ir automātisks paziņojums. / This is an automated notification.</p>
    </div>
  `;
  
  GmailApp.sendEmail(HR_EMAIL, subject, "", { htmlBody: htmlBody });
}
