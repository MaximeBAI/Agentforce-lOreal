// =========================================================
// 1. Configuration
// =========================================================
const CLIENT_ID = "CLIENT_ID";      
const CLIENT_SECRET = "CLIENT_SECRET"; 
const INSTANCE_DOMAIN = "INSTANCE_DOMAIN"; 
const AGENT_ID = 'AGENT_ID'; 
const API_BASE = 'https://api.salesforce.com/einstein/ai-agent/v1';

// ADJUST THIS: If you get the error of Google API limitation, lower this number to 3 or 5
const BATCH_SIZE = 50; 

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Testing Tools')
      .addItem('🚀 Run UAT Test', 'runParallelMultiRoundTest')
      .addToUi();
}

// =========================================================
// 2. Parallel Multi-Round Execution Logic with Batching
// =========================================================

function runParallelMultiRoundTest() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  const langColInput = ui.prompt("Setup", "Language Column letter (e.g., A):", ui.ButtonSet.OK_CANCEL);
  const uttColsInput = ui.prompt("Setup", "Utterance Column letters (e.g., B, C, D):", ui.ButtonSet.OK_CANCEL);
  if (langColInput.getSelectedButton() !== ui.Button.OK) return;

  const langCol = columnLetterToNumber(langColInput.getResponseText());
  const uttCols = uttColsInput.getResponseText().split(',').map(s => columnLetterToNumber(s.trim()));

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; 
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const token = getAccessToken();
  if (!token) { ui.alert("Auth Failed"); return; }

  // --- PHASE 1: Parallel Session Initiation in Batches ---
  console.log("TRACE: Initiating sessions in batches...");
  let activeSessions = [];
  
  for (let i = 0; i < dataRange.length; i += BATCH_SIZE) {
    const chunk = dataRange.slice(i, i + BATCH_SIZE);
    const sessionReqs = chunk.map(row => {
      const lang = row[langCol - 1];
      return lang ? buildStartSessionReq(token, lang) : null;
    }).filter(r => r !== null);

    const sessionResps = UrlFetchApp.fetchAll(sessionReqs);
    
    const chunkSessions = chunk.map((row, idx) => {
      const resp = sessionResps[idx];
      if (resp && (resp.getResponseCode() === 201 || resp.getResponseCode() === 200)) {
        const data = JSON.parse(resp.getContentText());
        return {
          rowNum: i + idx + 2,
          lang: row[langCol - 1],
          sessionId: data.sessionId,
          greeting: data.messages ? data.messages[0].message : "No Greeting",
          turns: [],
          isValid: true
        };
      }
      return { isValid: false };
    });
    activeSessions = activeSessions.concat(chunkSessions);
    Utilities.sleep(200); // Small pause to prevent bandwidth spikes
  }

  // --- PHASE 2: Sequential Rounds, Parallel Messages in Batches ---
  uttCols.forEach((colIndex, roundIdx) => {
    console.log(`TRACE: Round ${roundIdx + 1} processing...`);
    
    // We process each round in batches to satisfy the Google quota
    for (let i = 0; i < activeSessions.length; i += BATCH_SIZE) {
      const sessionChunk = activeSessions.slice(i, i + BATCH_SIZE);
      const dataChunk = dataRange.slice(i, i + BATCH_SIZE);
      
      const roundReqs = sessionChunk.map((session, idx) => {
        const userText = dataChunk[idx][colIndex - 1];
        return (session.isValid && userText) ? buildSendMessageReq(token, session.sessionId, userText) : null;
      }).filter(r => r !== null);

      if (roundReqs.length > 0) {
        const roundResps = UrlFetchApp.fetchAll(roundReqs);
        let resCount = 0;

        sessionChunk.forEach((session, idx) => {
          const userText = dataChunk[idx][colIndex - 1];
          if (session.isValid && userText) {
            const res = roundResps[resCount++];
            const responseBody = JSON.parse(res.getContentText());
            const agentMsg = responseBody.messages ? responseBody.messages.map(m => m.message).join("\n") : "Error";
            session.turns.push({ user: userText, agent: agentMsg });
          }
        });
      }
      Utilities.sleep(200);
    }
  });

  // --- PHASE 3: Logging to Timestamped Sheet ---
  const timestamp = Utilities.formatDate(new Date(), "GMT+1", "yyyyMMdd_HHmmss");
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(`Test_${timestamp}`);
  
  let headers = ["Original Row", "Lang", "Session ID", "Initial Greeting"];
  uttCols.forEach((_, i) => { headers.push(`Turn ${i+1}: User`, `Turn ${i+1}: Agent`); });
  logSheet.appendRow(headers);
  logSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d32f2f").setFontColor("white");

  activeSessions.forEach(s => {
    if (s.isValid) {
      let rowData = [s.rowNum, s.lang, s.sessionId, s.greeting];
      s.turns.forEach(t => { rowData.push(t.user, t.agent); });
      logSheet.appendRow(rowData);
      sheet.getRange(s.rowNum, langCol).setBackground("#D1FFD1");
      buildEndSessionReq(token, s.sessionId); 
    }
  });

  ui.alert(`Test complete. Results in: Test_${timestamp}`);
}

// =========================================================
// Helpers (Builders, Auth, Logic)
// =========================================================

function buildStartSessionReq(token, lang) {
  return {
    url: `${API_BASE}/agents/${AGENT_ID}/sessions`,
    method: 'post',
    contentType: 'application/json',
    headers: { "Authorization": "Bearer " + token },
    payload: JSON.stringify({
      "externalSessionKey": Utilities.getUuid(),
      "instanceConfig": { "endpoint": `https://${INSTANCE_DOMAIN}/` },
      "tz": "America/Los_Angeles",
      "variables": [{ "name": "$Context.EndUserLanguage", "type": "Text", "value": lang }],
      "featureSupport": "Streaming",
      "streamingCapabilities": { "chunkTypes": ["Text"] },
      "bypassUser": true
    }),
    muteHttpExceptions: true
  };
}

function buildSendMessageReq(token, sessionId, text) {
  return {
    url: `${API_BASE}/sessions/${sessionId}/messages`,
    method: 'post',
    contentType: 'application/json',
    headers: { "Authorization": "Bearer " + token, "Accept": "application/json" },
    payload: JSON.stringify({
      "message": { "sequenceId": Math.floor(Date.now() / 1000), "type": "Text", "text": text },
      "variables": []
    }),
    muteHttpExceptions: true
  };
}

function getAccessToken() {
  const url = `https://${INSTANCE_DOMAIN}/services/oauth2/token`;
  const headers = { "Authorization": "Basic " + Utilities.base64Encode(CLIENT_ID + ":" + CLIENT_SECRET) };
  const res = UrlFetchApp.fetch(url, { method: 'post', headers: headers, payload: { "grant_type": "client_credentials" }, muteHttpExceptions: true });
  return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()).access_token : null;
}

function buildEndSessionReq(token, sessionId) {
  UrlFetchApp.fetch(`${API_BASE}/sessions/${sessionId}`, {
    method: 'delete',
    headers: { "Authorization": "Bearer " + token, "x-session-end-reason": "UserRequest" },
    muteHttpExceptions: true
  });
}

function columnLetterToNumber(letter) {
  let column = 0, length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.toUpperCase().charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}
