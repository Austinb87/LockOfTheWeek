/**
 * LOCK OF THE WEEK — Google Sheets backend
 * ------------------------------------------------
 * SETUP (one time, ~5 minutes):
 * 1. Go to sheets.google.com and create a new blank spreadsheet.
 *    Name it whatever you want, e.g. "Lock of the Week".
 * 2. In the sheet, add a header row in row 1:
 *    Timestamp | Week Start | Week Label | Player | Pick | Odds | Status
 * 3. Extensions > Apps Script. Delete any starter code and paste in
 *    this entire file.
 * 4. Edit the PLAYERS array below to your 4 guys' real names
 *    (exact spelling — this is how the app matches people up).
 * 5. Click Deploy > New deployment.
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    Click Deploy, then authorize the permissions it asks for.
 * 6. Copy the "Web app URL" it gives you — you'll paste that into
 *    the CONFIG.WEBAPP_URL constant in the HTML app.
 * 7. Whenever you edit this script again, you must go to
 *    Deploy > Manage deployments > edit (pencil) > New version > Deploy
 *    for changes to actually go live.
 *
 * You are the only one with edit access to the Sheet or this script —
 * the web app can only ever append a row, never read or edit picks.
 */

// ---- EDIT THIS with your 4 guys' names, exactly as you want them shown ----
const PLAYERS = ['Austin', 'Baroni', 'Dobby', 'Dylan'];

// If your locks should reset on a day/time other than Thursday 12:00am,
// change this. 4 = Thursday (0=Sun, 1=Mon, 2=Tue, 3=Wed, 4=Thu, 5=Fri, 6=Sat)
const WEEK_START_DAY = 4;

function doGet(e) {
  try {
    const action = e.parameter.action || 'status';
    if (action === 'status') {
      return respond(getWeekStatus(e.parameter.week));
    }
    if (action === 'verifyAdmin') {
      const password = e.parameter.password || '';
      return respond({ success: verifyAdmin(password) });
    }
    return respond({ success: false, error: 'unknown action' });
  } catch (err) {
    return respond({ success: false, error: 'Server error: ' + err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // admin batch status update
    if (body.action === 'updateStatuses') {
      return updateStatuses(body);
    }

    const player = (body.player || '').trim();
    const pick = (body.pick || '').trim();
    const odds = (body.odds || '').trim();
    const weekId = (body.week || '').trim();

    if (!PLAYERS.includes(player)) {
      return respond({ success: false, error: 'Unrecognized player.' });
    }
    if (!pick) {
      return respond({ success: false, error: 'Pick cannot be empty.' });
    }
    if (!weekId) {
      return respond({ success: false, error: 'Missing week.' });
    }

    // Read sheet once
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    const data = sheet.getDataRange().getValues();

    // Enforce one lock per player per week - only check rows matching this week
    for (let i = 1; i < data.length; i++) {
      const rowWeekId = String(data[i][1]).trim();
      if (rowWeekId !== weekId) continue; // Skip rows from other weeks
      
      const rowPlayer = String(data[i][3]).trim();
      if (rowPlayer === player) {
        return respond({ success: false, error: 'Already submitted this week.' });
      }
    }

    // Append the row
    sheet.appendRow([
      new Date(),
      weekId,
      body.weekLabel || '',
      player,
      pick,
      odds
    ]);

    return respond({ success: true });
  } catch (err) {
    return respond({ success: false, error: 'Server error: ' + err.message });
  }
}

function verifyAdmin(password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet2");
    if (!sheet) return false;
    const stored = String(sheet.getRange(1, 1).getValue() || '');
    return stored === String(password);
  } catch (err) {
    return false;
  }
}

function updateStatuses(body) {
  try {
    const password = String(body.password || '');
    if (!verifyAdmin(password)) return respond({ success: false, error: 'Unauthorized' });

    const weekId = String(body.week || '');
    if (!weekId) return respond({ success: false, error: 'Missing week' });

    const updates = body.updates || [];
    if (!Array.isArray(updates) || updates.length === 0) {
      return respond({ success: false, error: 'No updates provided' });
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    const data = sheet.getDataRange().getValues();

    let updated = 0;
    updates.forEach(u => {
      const player = String(u.player || '').trim();
      const status = String(u.status || '').trim();
      if (!PLAYERS.includes(player)) return;
      if (!status) return;

      let found = false;
      for (let i = 1; i < data.length; i++) {
        const sheetWeek = formatWeekId(data[i][1]);
        if (sheetWeek !== weekId) continue; // Skip rows from other weeks
        
        const rowPlayer = String(data[i][3]).trim();
        if (rowPlayer === player) {
          // column G = 7
          sheet.getRange(i + 1, 7).setValue(status);
          found = true;
          updated++;
          break;
        }
      }

      if (!found) {
        // append a row with the status even if pick/odds are missing
        sheet.appendRow([new Date(), weekId, '', player, '', '', status]);
        updated++;
      }
    });

    return respond({ success: true, updated: updated });
  } catch (err) {
    return respond({ success: false, error: 'Server error: ' + err.message });
  }
}

function computeRecord(data, player) {
  let wins = 0, losses = 0, ties = 0;

  for (let i = 1; i < data.length; i++) {
    const rowPlayer = String(data[i][3]).trim();
    const result = String(data[i][6] || "");

    if (rowPlayer === player) {
      if (result === "✅") wins++;
      else if (result === "❌") losses++;
      else if (result === "👔") ties++;
    }
  }

  return { wins, losses, ties };
}

function formatWeekId(dateValue) {
  try {
    if (!dateValue) return '';
    const d = new Date(dateValue);
    return Utilities.formatDate(d, "GMT", "yyyy-MM-dd");
  } catch (err) {
    return String(dateValue).trim();
  }
}

function getWeekStatus(weekId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    const data = sheet.getDataRange().getValues();
    const submitted = {};
    PLAYERS.forEach(p => (submitted[p] = false));

    for (let i = 1; i < data.length; i++) {
      const sheetWeek = formatWeekId(data[i][1]);
      if (sheetWeek !== weekId) continue; // Skip rows from other weeks
      
      const player = String(data[i][3]).trim();

      if (submitted.hasOwnProperty(player)) {
        submitted[player] = {
          pick: data[i][4],   // Pick column
          odds: data[i][5],    // Odds column
          status: data[i][6]  // Status column
        };
      }
    }

    return {
      success: true,
      players: PLAYERS,
      submitted: submitted,
      records: PLAYERS.reduce((acc, p) => {
        acc[p] = computeRecord(data, p);
        return acc;
      }, {})
    };
  } catch (err) {
    return {
      success: false,
      error: 'Error fetching week status: ' + err.message
    };
  }
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
