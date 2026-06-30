/**
 * -----------------------------------------------------------------------
 * BRIEF PORTAL — SERVER-SIDE CODE (Add this to your Apps Script project)
 * -----------------------------------------------------------------------
 * Paste this into a NEW script file called "BriefPortal" in your
 * existing Apps Script project.
 * -----------------------------------------------------------------------
 */

/**
 * Returns average Brief Score per Client Champion for the leaderboard.
 * Column H (8) = Client Champion name, Column S (19) = Brief Score.
 */
function getLeaderboardData() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('2025/2026 Work List') || ss.getSheets()[0];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var numRows = lastRow - 1;

    // Read the two columns separately — avoids any indexing confusion
    var names  = sheet.getRange(2, 8,  numRows, 1).getValues(); // Column H
    var scores = sheet.getRange(2, 19, numRows, 1).getValues(); // Column S

    // key = normalised (lowercase, no spaces) for deduplication
    // display = first seen version of the name (proper casing)
    var totals  = {};
    var display = {};
    for (var i = 0; i < numRows; i++) {
      var raw   = String(names[i][0]  || '').trim();
      var score = Number(scores[i][0]);
      if (!raw || !score || isNaN(score)) continue;
      var key = raw.toLowerCase().replace(/\s+/g, ''); // "yik wen" → "yikwen"
      if (!totals[key]) {
        totals[key]  = { total: 0, count: 0 };
        display[key] = raw; // keep first-seen casing as display name
      }
      totals[key].total += score;
      totals[key].count++;
    }

    Logger.log('Totals: ' + JSON.stringify(totals));

    // ── Bayesian average (same method as IMDb/Letterboxd) ──
    // Pulls low-count scores toward the group mean, so 1 brief at 8.0
    // doesn't outrank someone with 18 briefs at 6.1.
    // Formula: (C × globalMean + personTotal) / (C + personCount)
    // C = confidence constant — how many briefs it takes to "trust" a score fully.
    var C = 5;
    var grandTotal = 0, grandCount = 0;
    for (var k in totals) { grandTotal += totals[k].total; grandCount += totals[k].count; }
    var globalMean = grandCount > 0 ? grandTotal / grandCount : 6;

    Logger.log('Global mean: ' + globalMean.toFixed(2) + ' across ' + grandCount + ' briefs');

    var result = [];
    for (var k in totals) {
      var bayesian = (C * globalMean + totals[k].total) / (C + totals[k].count);
      result.push({
        name:  display[k],
        avg:   Math.round(bayesian * 10) / 10,
        raw:   Math.round((totals[k].total / totals[k].count) * 10) / 10,
        count: totals[k].count
      });
    }
    result.sort(function(a, b) { return b.avg - a.avg || b.count - a.count; });

    Logger.log('Result: ' + result.length + ' people');
    return result;

  } catch(e) {
    Logger.log('getLeaderboardData error: ' + e.toString());
    return [];
  }
}

/**
 * Serves the Brief Portal or Strategist Dashboard based on ?page= parameter.
 *
 * DEPLOYMENT NOTE:
 *   - Brief Portal:          Deploy with "Execute as: Me"    → share URL to Sales
 *   - Strategist Dashboard:  Deploy with "Execute as: User accessing the web app"
 *                            → share URL + ?page=strategist to Strategists
 *     The "User" setting is required so Session.getActiveUser().getEmail()
 *     returns the logged-in strategist's email (not the owner's).
 */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'brief';
  if (page === 'strategist') {
    return HtmlService.createHtmlOutputFromFile('StrategistDashboard')
      .setTitle('REV Media — My Briefs')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'brandintel') {
    return HtmlService.createHtmlOutputFromFile('BrandIntel')
      .setTitle('REV Media — Brand Intel')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('BriefPortal')
    .setTitle('REV Media — Brief Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Main handler called when Sales submits the form.
 * Creates a properly-named folder in the intake Drive folder,
 * then places the brief content (Google Doc or Quill text) inside it.
 */
function handleBriefSubmission(formData) {
  try {
    var intakeFolder = DriveApp.getFolderById(CONFIG.newFilesFolderId);

    // ── KLR path ──────────────────────────────────────────────
    if (formData.requestType === 'klr') {
      var klrFolderName = [
        formData.salesName,
        formData.stratName,
        'KLR',
        (formData.brand || '').replace(/_/g, '-'),
        formData.deadline
      ].join('_');

      var klrFolder = intakeFolder.createFolder(klrFolderName);

      // ── Backend loggers: inspect what was received ──
      Logger.log('KLR formData.klrImages count: ' + ((formData.klrImages || []).length));
      Logger.log('KLR formData.questions length: ' + (formData.questions || '').length);
      Logger.log('KLR questions first 500 chars: ' + (formData.questions || '').substring(0, 500));

      // ── Backend fallback: extract data: URL images from questions HTML ──
      // If the frontend already ran extractKlrImages(), formData.klrImages is populated
      // and formData.questions has [📷 Image N] placeholders.
      // If an old frontend is used, the data: URLs are still in the HTML — extract them here.
      var klrImages = (formData.klrImages && formData.klrImages.length > 0)
        ? formData.klrImages
        : [];
      if (klrImages.length === 0 && formData.questions) {
        var imgCounter_ = 0;
        formData.questions = formData.questions.replace(
          /<img[^>]+src="(data:([^;]+);base64,([^"]+))"[^>]*\/?>/gi,
          function(match, src, mimeType, b64) {
            imgCounter_++;
            klrImages.push({ data: b64, mimeType: mimeType || 'image/png', index: imgCounter_ });
            return '<span>##IMG_' + imgCounter_ + '##</span>';
          }
        );
        if (imgCounter_ > 0) Logger.log('Backend extracted ' + imgCounter_ + ' image(s) from questions HTML.');
      }

      // ── Upload PCR files BEFORE creating the doc so their URLs can be embedded ──
      // Col O = PCR only. First PCR file/link → Col O. Extra PCR files → embedded in doc.
      var pcrColOFormula = '';   // will go into fileLinksArr[1] → sheet Col O
      var pcrExtraLinks  = [];   // additional PCR files (rare) — embedded in doc

      if (formData.pcrType === 'link' && formData.pcrLink) {
        // Paste Google Drive/Slides link path
        pcrColOFormula = '=HYPERLINK("' + formData.pcrLink + '", "📂 Open PCR")';
        try { handleGoogleDocLink_(formData.pcrLink, klrFolderName + ' - PCR', klrFolder); } catch(e) {}
      } else if (formData.pcrAttachments && formData.pcrAttachments.length > 0) {
        // File upload path — first file → Col O, extras → embed in doc
        formData.pcrAttachments.forEach(function(file, idx) {
          try {
            var decoded  = Utilities.base64Decode(file.data);
            var blob     = Utilities.newBlob(decoded, file.mimeType || 'application/octet-stream', file.name);
            var uploaded = klrFolder.createFile(blob);
            if (idx === 0) {
              pcrColOFormula = '=HYPERLINK("' + uploaded.getUrl() + '", "📂 Open PCR")';
            } else {
              pcrExtraLinks.push({ name: file.name, url: uploaded.getUrl() });
            }
          } catch(e) { Logger.log('KLR PCR attachment failed: ' + (file.name || '') + ' — ' + e.message); }
        });
      }

      // ── Upload supporting attachments → embed Drive links in the doc (NOT in sheet cols) ──
      var suppLinks = [];
      if (formData.attachments && formData.attachments.length > 0) {
        formData.attachments.forEach(function(file) {
          try {
            var decoded  = Utilities.base64Decode(file.data);
            var blob     = Utilities.newBlob(decoded, file.mimeType || 'application/octet-stream', file.name);
            var uploaded = klrFolder.createFile(blob);
            suppLinks.push({ name: file.name, url: uploaded.getUrl() });
            Logger.log('KLR supporting file uploaded: ' + file.name);
          } catch(e) { Logger.log('KLR attachment upload failed: ' + (file.name || '') + ' — ' + e.message); }
        });
      }

      // ── Build PCR section for the KLR Brief doc ──
      var pcrSection;
      if (formData.noPcr) {
        pcrSection = '<h2>PCR Status</h2>' +
          '<p><em>Campaign still in progress — PCR will be shared once it wraps up.</em></p>';
      } else if (formData.pcrType === 'link' && formData.pcrLink) {
        pcrSection = '<h2>PCR</h2><p><a href="' + formData.pcrLink + '">' + formData.pcrLink + '</a></p>';
      } else if (pcrColOFormula) {
        // File uploaded — note it in the doc; extras embedded below
        pcrSection = '<h2>PCR</h2><p>PCR file attached — see Drive folder link in sheet Col O.</p>';
        if (pcrExtraLinks.length > 0) {
          pcrSection += '<ul>' + pcrExtraLinks.map(function(f) {
            return '<li><a href="' + f.url + '">' + f.name + '</a></li>';
          }).join('') + '</ul>';
        }
      } else {
        pcrSection = ''; // noPcr toggle not used, nothing provided (shouldn't normally reach here)
      }

      // ── Build Supporting Files section for the KLR Brief doc ──
      var suppSection = '';
      if (suppLinks.length > 0) {
        suppSection = '<h2>Supporting Files</h2><ul>' +
          suppLinks.map(function(f) {
            return '<li><a href="' + f.url + '">' + f.name + '</a></li>';
          }).join('') + '</ul>';
      }

      // ── Assemble and create the KLR Google Doc ──
      var klrHtml =
        '<h1>KLR Request — ' + formData.brand + '</h1>' +
        '<table>' +
          '<tr><td><b>Requested by</b></td><td>' + formData.salesName + '</td></tr>' +
          '<tr><td><b>Strategist</b></td><td>'   + formData.stratName  + '</td></tr>' +
          '<tr><td><b>Client / Brand</b></td><td>' + formData.brand    + '</td></tr>' +
          '<tr><td><b>Campaign Name</b></td><td>' + formData.campaignName + '</td></tr>' +
          '<tr><td><b>Deadline</b></td><td>'     + formData.deadline   + '</td></tr>' +
        '</table>' +
        "<h2>Client's Specific Questions</h2>" +
        (formData.questions && formData.questions.trim() ? formData.questions : '<p>None provided.</p>') +
        pcrSection +
        suppSection;

      var klrDocBlob = Utilities.newBlob(klrHtml, 'text/html', 'KLR Request Details');
      var klrDocResource = {
        title: 'KLR Request Details — ' + formData.brand,
        mimeType: 'application/vnd.google-apps.document',
        parents: [{ id: klrFolder.getId() }]
      };
      // Create KLR Google Doc and capture its URL for sheet links
      var klrDocFile = Drive.Files.insert(klrDocResource, klrDocBlob, { convert: true });
      var klrDocUrl  = 'https://docs.google.com/document/d/' + klrDocFile.id + '/edit';
      var klrDocHyperlink = '=HYPERLINK("' + klrDocUrl + '", "📄 View KLR Brief")';
      Logger.log('KLR Google Doc created: ' + klrDocUrl);

      // ── Insert images inline — find ##IMG_N## markers and replace with actual images ──
      if (klrImages.length > 0) {
        try {
          var imgDoc  = DocumentApp.openById(klrDocFile.id);
          var imgBody = imgDoc.getBody();
          klrImages.forEach(function(img) {
            var marker = '##IMG_' + img.index + '##';
            var found  = imgBody.findText(marker);
            var imgBlob = Utilities.newBlob(
              Utilities.base64Decode(img.data),
              img.mimeType || 'image/png',
              'image_' + img.index
            );
            if (found) {
              var para = found.getElement().getParent();
              var idx  = imgBody.getChildIndex(para);
              imgBody.insertImage(idx, imgBlob);
              imgBody.removeChild(para);
              Logger.log('KLR: image ' + img.index + ' inserted inline at marker position.');
            } else {
              // Marker not found (e.g. backend fallback path) — append at end
              imgBody.appendImage(imgBlob);
              Logger.log('KLR: image ' + img.index + ' marker not found, appended at end.');
            }
          });
          imgDoc.saveAndClose();
          Logger.log('Inserted ' + klrImages.length + ' image(s) into KLR doc.');
        } catch(ie) {
          Logger.log('KLR image insertion failed (non-fatal): ' + ie.toString());
        }
      }

      // ── Sheet column layout: N=KLR doc, O=PCR only, P/Q/R=reserved for updates ──
      var fileLinksArr = ['', '', '', '', ''];
      fileLinksArr[0] = klrDocHyperlink;   // Col N: KLR Brief Doc
      fileLinksArr[1] = pcrColOFormula;    // Col O: PCR only (blank if noPcr toggled)
      // fileLinksArr[2-4] stay empty — reserved for update docs in P, Q, R

      // ── Write full row to Work List (matching Sales brief format, A–W) ──
      try {
        var wlSs    = SpreadsheetApp.getActiveSpreadsheet();
        var wlSheet = wlSs.getSheetByName(CONFIG.sheetName) || wlSs.getSheets()[0];

        // formData.deadline is dd/MM/yyyy (converted on frontend)
        var dlParts    = (formData.deadline || '').split('/');
        var klrDeadline = dlParts.length === 3
          ? new Date(parseInt(dlParts[2]), parseInt(dlParts[1]) - 1, parseInt(dlParts[0]))
          : '';

        var trackingCode    = generateTrackingCode();
        var lastRow         = wlSheet.getLastRow() + 1;
        var slaDateFormula  = '=WORKDAY(U' + lastRow + ', 5)';

        // Gemini industry classification — uses brand, campaign name, and client context
        var klrIndustry = classifyKlrIndustry_(
          formData.brand        || '',
          formData.campaignName || '',
          formData.questions    || ''
        );

        var klrRow = [
          trackingCode,                                // A: Tracking Code
          'KLR',                                       // B: Type of Deck
          'N/A',                                       // C: Parent Company
          formData.brand || '',                        // D: Brand
          'KLR For ' + (formData.brand || ''),         // E: Campaign / Project Name
          0,                                           // F: Budget (RM 0.00)
          formData.stratName || '',                    // G: Strat
          formData.salesName || '',                    // H: Client Champion (requester)
          klrIndustry,                                 // I: Industry (Gemini-classified)
          'Pending',                                   // J: Status
          '',                                          // K: Proposal Link
          klrDeadline,                                 // L: Deadline
          '',                                          // M: Summary Doc Link (left empty — brief is in Col N)
          fileLinksArr[0],                             // N: OG Brief Link 1 (KLR doc)
          fileLinksArr[1],                             // O: OG Brief Link 2 (PCR)
          fileLinksArr[2],                             // P: OG Brief Link 3
          fileLinksArr[3],                             // Q: OG Brief Link 4
          fileLinksArr[4],                             // R: OG Brief Link 5
          '',                                          // S: Brief Score (N/A for KLR)
          'N/A',                                       // T: Missing Info
          new Date(),                                  // U: Generated On
          5,                                           // V: SLA Days
          slaDateFormula                               // W: SLA Target Date (=WORKDAY(U{row}, 5))
        ];

        wlSheet.appendRow(klrRow);
        wlSheet.getRange(lastRow, 6).setNumberFormat('"RM" #,##0.00');  // F: Budget
        if (klrDeadline) wlSheet.getRange(lastRow, 12).setNumberFormat('dd/mm/yyyy'); // L: Deadline
        wlSheet.getRange(lastRow, 23).setNumberFormat('dd/mm/yyyy');    // W: SLA Target Date
        Logger.log('KLR row added to Work List: ' + trackingCode + ' — ' + klrFolderName);
      } catch(wlErr) {
        Logger.log('KLR Work List row failed (non-fatal): ' + wlErr.toString());
      }

      // ── Notify Campaign + Strategy Chat space ──
      try {
        var pcrStatusLine = formData.noPcr
          ? '*PCR:* _Campaign still running — will be shared once it wraps up_'
          : '*PCR:* Attached';
        var inboundMsg =
          '📋 *New KLR Request*\n\n' +
          '*Brand:* '        + (formData.brand        || 'N/A') + '\n' +
          '*Campaign:* '     + (formData.campaignName || 'N/A') + '\n' +
          '*Requested by:* ' + (formData.salesName    || 'N/A') + '\n' +
          '*Assigned to:* '  + getMention(formData.stratName)   + '\n' +
          '*Deadline:* '     + (formData.deadline     || 'N/A') + '\n' +
          pcrStatusLine;
        sendCampaignWebhook_(inboundMsg);
      } catch(whErr) {
        Logger.log('KLR inbound webhook failed (non-fatal): ' + whErr.toString());
      }

      Logger.log('KLR request submitted: ' + klrFolderName);
      return { success: true, folderName: klrFolderName };
    }

    // ── Proposal path (existing) ───────────────────────────────
    // 1. Build the folder name using the same naming convention the pipeline expects:
    //    SalesName_StratName_Budget_DeckType_DD/MM/YYYY
    var folderName = [
      formData.salesName,
      formData.stratName,
      formData.budget,
      formData.deckType,
      formData.deadline
    ].join('_');

    // 2. Create the folder inside the intake folder
    var newFolder = intakeFolder.createFolder(folderName);

    // 3. Handle the brief content
    if (formData.briefType === 'link' && formData.googleDocUrl) {
      handleGoogleDocLink_(formData.googleDocUrl, folderName, newFolder);
    } else if (formData.briefType === 'editor' && formData.quillHtml) {
      handleQuillContent_(formData.quillHtml, folderName, newFolder);
    }

    // 4. Save any attached files into the folder
    if (formData.attachments && formData.attachments.length > 0) {
      formData.attachments.forEach(function(file) {
        try {
          var decoded = Utilities.base64Decode(file.data);
          var blob = Utilities.newBlob(decoded, file.mimeType || 'application/octet-stream', file.name);
          newFolder.createFile(blob);
          Logger.log('Attachment saved: ' + file.name);
        } catch (fileErr) {
          Logger.log('Failed to save attachment ' + file.name + ': ' + fileErr.message);
        }
      });
    }

    Logger.log('Brief submitted successfully: ' + folderName);
    return { success: true, folderName: folderName };

  } catch (e) {
    Logger.log('Brief submission error: ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * Posts a message to the joint Campaign + Strategy Google Chat space.
 * Webhook URL stored in Script Properties as CAMPAIGN_WEBHOOK_URL.
 * Fails silently if the property is not yet set.
 */
function sendCampaignWebhook_(message) {
  var url = PropertiesService.getScriptProperties().getProperty('CAMPAIGN_WEBHOOK_URL');
  if (!url) {
    Logger.log('sendCampaignWebhook_: CAMPAIGN_WEBHOOK_URL not set — skipping notification.');
    return;
  }
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message }),
      muteHttpExceptions: true
    });
  } catch(e) {
    Logger.log('sendCampaignWebhook_ error: ' + e.toString());
  }
}

/**
 * Handles the case where Sales pastes a Google Doc link.
 * Tries to copy the actual doc into the folder.
 * Falls back to saving the link in a text file if access is denied.
 */
function handleGoogleDocLink_(url, folderName, targetFolder) {
  var docId = extractGoogleDocId_(url);

  if (!docId) {
    // Not a valid Google Doc URL — save the link as a text file
    var blob = Utilities.newBlob('Brief Link: ' + url, 'text/plain', 'Brief Link.txt');
    targetFolder.createFile(blob);
    return;
  }

  try {
    // Try to copy the Google Doc into the folder
    var sourceFile = DriveApp.getFileById(docId);
    sourceFile.makeCopy(folderName + ' - Brief', targetFolder);
    Logger.log('Google Doc copied successfully: ' + docId);
  } catch (e) {
    Logger.log('Could not copy Google Doc (access issue): ' + e.message + '. Saving link instead.');
    // Fallback: save the link as a text file so the folder still has something
    var blob = Utilities.newBlob(
      'Brief Document Link:\n' + url + '\n\nNote: The script could not automatically copy this document.\nPlease ensure it is shared with the REV Media team.',
      'text/plain',
      'Brief Link.txt'
    );
    targetFolder.createFile(blob);
  }
}

/**
 * Handles the case where Sales writes in the editor.
 * Uses the Drive API to convert HTML → Google Doc, preserving all formatting
 * (bold, italic, underline, bullet lists, numbered lists, links).
 */
function handleQuillContent_(html, folderName, targetFolder) {
  if (!html || html.replace(/<[^>]*>/g, '').trim().length < 5) {
    Logger.log('Editor content was empty.');
    return;
  }

  try {
    // Drive API converts HTML to a native Google Doc with formatting intact
    var htmlBlob = Utilities.newBlob(html, 'text/html', folderName + ' - Brief');
    var resource = {
      title: folderName + ' - Brief',
      mimeType: 'application/vnd.google-apps.document',
      parents: [{ id: targetFolder.getId() }]
    };
    Drive.Files.insert(resource, htmlBlob, { convert: true });
    Logger.log('Google Doc created with formatting preserved via Drive API.');

  } catch (e) {
    Logger.log('Drive API HTML conversion failed, falling back to plain text: ' + e.message);
    // Fallback: plain text if Drive API fails for any reason
    var plainText = htmlToPlainText_(html);
    var doc = DocumentApp.create(folderName + ' - Brief');
    doc.getBody().setText(plainText);
    doc.saveAndClose();
    var docFile = DriveApp.getFileById(doc.getId());
    targetFolder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile);
  }
}

/**
 * Extracts the document ID from a Google Docs URL.
 * Handles all common URL formats.
 */
function extractGoogleDocId_(url) {
  var match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

// -----------------------------------------------------------------------
// STRATEGIST DASHBOARD — BACKEND FUNCTIONS
// -----------------------------------------------------------------------

/**
 * Returns the logged-in strategist's name and their filtered brief rows.
 * Matches Session.getActiveUser().getEmail() against STRAT_EMAILS (from main script).
 * REQUIRES: web app deployed as "Execute as: User accessing the web app".
 */
function getStrategistData(completedOnly) {
  try {
    var currentEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
    if (!currentEmail) {
      return { found: false, error: 'no_email', email: '' };
    }

    // Reverse-lookup: find which strat name maps to this email
    var stratName = null;
    for (var name in STRAT_EMAILS) {
      if (STRAT_EMAILS[name].toLowerCase().trim() === currentEmail) {
        stratName = name;
        break;
      }
    }

    if (!stratName) {
      return { found: false, error: 'not_found', email: currentEmail };
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetName) || ss.getSheets()[0];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { found: true, name: stratName, briefs: [] };

    var numRows = lastRow - 1;

    // Values for non-formula columns (A-Z = cols 1-26, includes Brief Summary in col Z)
    var values = sheet.getRange(2, 1, numRows, 26).getValues();
    // Formulas for Summary + Brief link columns (M-R = cols 13-18, 6 wide)
    var linkFormulas = sheet.getRange(2, 13, numRows, 6).getFormulas();

    // Dan and Danesh are the same person — handle both
    var nameVariants = [stratName];
    if (stratName === 'Danesh') nameVariants.push('Dan');
    if (stratName === 'Dan')    nameVariants.push('Danesh');

    var briefs = [];
    // Read all Script Properties once so we can check KLR update flags in the loop
    var klrUpdateProps = PropertiesService.getScriptProperties().getProperties();

    for (var i = 0; i < numRows; i++) {
      var row       = values[i];
      var rowStrat  = String(row[CONFIG.colIndices.strat - 1] || '').trim();
      if (nameVariants.indexOf(rowStrat) === -1) continue;

      // Deadline
      var dl = row[CONFIG.colIndices.deadline - 1];
      var deadlineStr = 'N/A';
      var deadlineMs  = null;
      if (dl instanceof Date) {
        deadlineStr = Utilities.formatDate(dl, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        deadlineMs  = dl.getTime();
      }

      // Budget (Column F = index 5)
      var budgetRaw = String(row[5] || '').replace(/[^0-9.]/g, '');
      var budgetNum = parseFloat(budgetRaw);
      var budget    = (!isNaN(budgetNum) && budgetNum > 0)
        ? 'RM ' + budgetNum.toLocaleString('en-MY') : '';

      // Summary doc (Column M = linkFormulas[i][0])
      var summaryUrl = extractHyperlinkUrl_(linkFormulas[i][0]) ||
                       extractHyperlinkUrl_(String(row[12] || ''));

      // Brief links (Columns N-R = linkFormulas[i][1..5])
      var briefLinks = [];
      for (var li = 1; li <= 5; li++) {
        var url = extractHyperlinkUrl_(linkFormulas[i][li]) ||
                  extractHyperlinkUrl_(String(row[12 + li] || ''));
        if (url) briefLinks.push(url);
      }

      // KLR "new update" flag — stored in Script Properties by updateKlrContext
      var deckType_  = String(row[1] || '').trim();
      var hasUpdate  = (deckType_ === 'KLR') &&
                       !!(klrUpdateProps['KLR_UPDATE_' + (i + 2)]);

      // KLR requester Chat ID — for "Ask More in GChat" deep link on Strategist Dashboard
      var requesterName_  = String(row[7] || '').trim(); // Col H: requester / Client Champion
      var requesterChatId = (USER_DIRECTORY && USER_DIRECTORY[requesterName_]) || '';

      briefs.push({
        row:             i + 2,           // 1-based sheet row for updates
        brand:           String(row[CONFIG.colIndices.brand - 1]        || ''),
        project:         String(row[CONFIG.colIndices.project - 1]      || ''),
        status:          String(row[CONFIG.colIndices.status - 1]       || ''),
        proposalLink:    String(row[CONFIG.colIndices.proposalLink - 1] || ''),
        cc:              String(row[CONFIG.colIndices.cc - 1]           || ''),
        deadline:        deadlineStr,
        deadlineMs:      deadlineMs,
        budget:          budget,
        score:           row[18] || '',   // Column S (index 18) = Brief Score
        deck:            String(row[1]  || '').trim(),  // Column B = Type of Deck
        industry:        String(row[8]  || '').trim(),  // Column I = Industry or Internal
        briefSummary:    String(row[25] || '').trim(),  // Column Z (index 25) = Brief Summary
        summaryUrl:      summaryUrl || '',
        briefLinks:      briefLinks,
        hasUpdate:       hasUpdate,
        requesterChatId: requesterChatId  // numeric Google Chat ID (for KLR "Ask More in GChat" button)
      });
    }

    // Sort: non-done by deadline ascending, done at the bottom
    briefs.sort(function(a, b) {
      var aDone = a.status.toLowerCase() === 'done';
      var bDone = b.status.toLowerCase() === 'done';
      if (aDone !== bDone) return aDone ? 1 : -1;
      // Pending: earliest deadline first; Done: most recent deadline first
      var mult = aDone ? -1 : 1;
      if (a.deadlineMs && b.deadlineMs) return mult * (a.deadlineMs - b.deadlineMs);
      if (a.deadlineMs) return -1;
      if (b.deadlineMs) return 1;
      return 0;
    });

    var pending = briefs.filter(function(b) { return b.status.toLowerCase() !== 'done'; });
    var done    = briefs.filter(function(b) { return b.status.toLowerCase() === 'done'; });

    if (completedOnly) {
      return { found: true, name: stratName, email: currentEmail, briefs: done, strats: Object.keys(STRAT_EMAILS).sort() };
    }

    return {
      found: true, name: stratName, email: currentEmail,
      briefs: pending,
      donePreview: done.slice(0, 5),
      doneTotal: done.length,
      strats: Object.keys(STRAT_EMAILS).sort()
    };

  } catch (e) {
    Logger.log('getStrategistData error: ' + e.toString());
    return { found: false, error: e.message, email: '' };
  }
}

/**
 * Writes proposal link (Column K) and marks status as Done (Column J).
 * Also fires the Google Chat notification directly — onEdit does NOT fire
 * on programmatic setValue() calls, so we can't rely on onStatusChange here.
 */
function submitProposal(rowIndex, proposalLink) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetName) || ss.getSheets()[0];

    if (proposalLink) {
      sheet.getRange(rowIndex, CONFIG.colIndices.proposalLink).setValue(proposalLink);
    }
    sheet.getRange(rowIndex, CONFIG.colIndices.status).setValue('Done');

    // Read row data for the notification
    var row     = sheet.getRange(rowIndex, 1, 1, CONFIG.colIndices.cc).getValues()[0];
    var deck    = String(row[1] || '').trim();   // Col B = Type of Deck
    var brand   = row[CONFIG.colIndices.brand   - 1];
    var project = row[CONFIG.colIndices.project - 1];
    var strat   = row[CONFIG.colIndices.strat   - 1];
    var cc      = row[CONFIG.colIndices.cc      - 1];

    if (deck === 'KLR') {
      // KLR done → notify the joint Campaign + Strategy space
      var klrDoneMsg =
        '✅ *KLR Deck Ready*\n\n' +
        '*Brand:* '        + brand + '\n' +
        '*Campaign:* '     + project + '\n' +
        '*Strat:* '        + getMention(strat) + ' has completed the KLR deck.\n' +
        '*Requested by:* ' + getMention(cc) + '\n' +
        '*Deck:* '         + (proposalLink || 'N/A');
      sendCampaignWebhook_(klrDoneMsg);
    } else {
      // Regular proposal → notify Strategy space as before
      var message = '✅ *Proposal Completed*\n\n' +
        '*Brand:* '    + brand + '\n' +
        '*Strat:* '    + getMention(strat) + ' has marked the project as done.\n' +
        '*Project:* '  + project + '\n' +
        '*CC:* '       + getMention(cc) + '\n' +
        '*Proposal:* ' + (proposalLink || 'N/A');
      sendWebhook(message);

      // Auto-index only for non-KLR proposals
      try { indexRowIfReady_(sheet, rowIndex); } catch(idxErr) {
        Logger.log('Auto-index warning (non-fatal): ' + idxErr.toString());
      }
    }

    return { success: true };
  } catch (e) {
    Logger.log('submitProposal error: ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * Updates the deadline (Column L) for a given row.
 * dateStr is dd/MM/yyyy as typed/selected by the strategist.
 */
function updateDeadline(rowIndex, dateStr) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetName) || ss.getSheets()[0];
    var parts = dateStr.split('/');
    if (parts.length !== 3) return { success: false, error: 'Invalid date format' };
    var newDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
    sheet.getRange(rowIndex, CONFIG.colIndices.deadline).setValue(newDate);
    sheet.getRange(rowIndex, CONFIG.colIndices.deadline).setNumberFormat('dd/mm/yyyy');

    // Notify relevant space based on deck type
    var rowData   = sheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
    var deckType  = String(rowData[1] || '').trim(); // Col B
    var brand     = String(rowData[3] || '').trim(); // Col D
    var campaign  = String(rowData[4] || '').trim(); // Col E
    var strat     = String(rowData[6] || '').trim(); // Col G
    var ccName    = String(rowData[7] || '').trim(); // Col H

    if (deckType === 'KLR') {
      var klrMsg =
        '📅 *KLR Deadline Proposed*\n\n' +
        '*Brand:* '        + brand    + '\n' +
        '*Campaign:* '     + campaign + '\n' +
        '*Strategist:* '   + getMention(strat) + ' has proposed a new deadline.\n' +
        '*New Deadline:* ' + dateStr  + '\n' +
        '*Requested by:* ' + getMention(ccName) + '\n\n' +
        '_Please confirm if this works for you._';
      sendCampaignWebhook_(klrMsg);
    } else {
      var msg =
        '📅 *Deadline Updated*\n\n' +
        '*Brand:* '        + brand    + '\n' +
        '*Project:* '      + campaign + '\n' +
        '*New Deadline:* ' + dateStr  + '\n' +
        '*CC:* '           + getMention(ccName) + '\n' +
        '*Strat:* '        + getMention(strat);
      sendWebhook(msg);
    }

    return { success: true };
  } catch (e) {
    Logger.log('updateDeadline error: ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * Creates a task in the logged-in user's Google Tasks default list.
 * Requires "Execute as: User accessing the web app" deployment.
 */
/**
 * Creates a task in the user's default Google Tasks list.
 * Requires: Apps Script > Services > Tasks API (must be enabled manually).
 */
function addToGoogleTasks(title, notes, deadlineStr) {
  try {
    var task = { title: title, notes: notes };
    // deadlineStr is dd/MM/yyyy — parse directly to avoid timezone shifts
    if (deadlineStr && deadlineStr !== 'N/A') {
      var parts = deadlineStr.split('/');
      if (parts.length === 3) {
        task.due = parts[2] + '-' + parts[1] + '-' + parts[0] + 'T00:00:00.000Z';
      }
    }
    Tasks.Tasks.insert(task, '@default');
    return { success: true };
  } catch (e) {
    Logger.log('addToGoogleTasks error: ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * Reassigns a brief (Column G = strat) to another strategist and fires a
 * Google Chat notification. Called when a strategist is going on leave.
 */
function handoffBrief(rowIndex, newStratName) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetName) || ss.getSheets()[0];

    var row     = sheet.getRange(rowIndex, 1, 1, CONFIG.colIndices.cc).getValues()[0];
    var brand   = row[CONFIG.colIndices.brand   - 1];
    var project = row[CONFIG.colIndices.project - 1];
    var oldStrat = row[CONFIG.colIndices.strat  - 1];

    sheet.getRange(rowIndex, CONFIG.colIndices.strat).setValue(newStratName);

    var message = '🔄 *Strat Update / Handover*\n\n' +
      '*Brand:* '  + brand + '\n' +
      '*Project:* ' + project + '\n' +
      '*Update:* ' + getMention(newStratName) + ' has taken over from ' + oldStrat + '.';
    sendWebhook(message);

    return { success: true };
  } catch (e) {
    Logger.log('handoffBrief error: ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * Extracts the URL from a HYPERLINK formula string, or returns a plain URL.
 * Handles: =HYPERLINK("url","text")  and  https://...
 */
function extractHyperlinkUrl_(cellStr) {
  if (!cellStr) return null;
  var s = String(cellStr).trim();
  var match = s.match(/HYPERLINK\("([^"]+)"/i);
  if (match) return match[1];
  if (s.indexOf('http') === 0) return s;
  return null;
}

// ═══════════════════════════════════════════════════════════════════
// KLR — INDUSTRY CLASSIFICATION
// ═══════════════════════════════════════════════════════════════════

/**
 * Uses Gemini to pick the closest industry from APPROVED_INDUSTRIES
 * based on the KLR form context: brand name, campaign name, and the
 * client questions / additional context (HTML stripped to plain text).
 * Falls back to 'N/A' if Gemini is unavailable or returns an unrecognised value.
 */
function classifyKlrIndustry_(brand, campaignName, questionsHtml) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) { Logger.log('classifyKlrIndustry_: no GEMINI_API_KEY'); return 'N/A'; }

    // Strip HTML tags so Gemini gets clean text
    var contextText = (questionsHtml || '')
      .replace(/<[^>]*>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .substring(0, 800); // cap tokens

    var input = 'Brand: ' + (brand || '(unknown)') +
                '\nCampaign: ' + (campaignName || '(unknown)');
    if (contextText) input += '\nClient context: ' + contextText;

    var prompt =
      'You are classifying a marketing campaign into a standard industry category.\n\n' +
      'Campaign details:\n' + input + '\n\n' +
      'Choose the SINGLE most appropriate industry from this approved list:\n' +
      APPROVED_INDUSTRIES + '\n\n' +
      'Rules:\n' +
      '- Reply with ONLY the exact industry name as it appears in the list.\n' +
      '- No explanation, no punctuation, just the industry name.\n' +
      '- If you are genuinely unsure, reply with: N/A';

    var resp = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.1, maxOutputTokens: 30 }
        }),
        muteHttpExceptions: true
      }
    );

    var json     = JSON.parse(resp.getContentText());
    var industry = ((json.candidates || [])[0] || {});
    industry     = (((industry.content || {}).parts || [])[0] || {}).text || '';
    industry     = industry.trim();

    // Validate — must be in the approved list
    var approved = APPROVED_INDUSTRIES.split(',').map(function(i) { return i.trim(); });
    if (approved.indexOf(industry) !== -1) return industry;

    // Fuzzy rescue: if the response contains an approved name as a substring
    for (var i = 0; i < approved.length; i++) {
      if (industry.toLowerCase().indexOf(approved[i].toLowerCase()) !== -1) return approved[i];
    }

    Logger.log('classifyKlrIndustry_: unrecognised response "' + industry + '" — using N/A');
    return 'N/A';

  } catch(e) {
    Logger.log('classifyKlrIndustry_ error: ' + e.toString());
    return 'N/A';
  }
}

// ═══════════════════════════════════════════════════════════════════
// KLR — UPDATE FLOW (portal search + additional context)
// ═══════════════════════════════════════════════════════════════════

/**
 * Returns all KLR rows in the Work List submitted by requesterName (Col H).
 * Called by the "Update KLR" portal panel.
 */
function getKlrSubmissions(requesterName) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetName);
    if (!sheet) return [];

    var data    = sheet.getDataRange().getValues();
    var tz      = Session.getScriptTimeZone();
    var results = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[1]).trim() !== 'KLR') continue;
      if (String(row[7]).trim() !== requesterName) continue;

      var dl = row[11]; // Col L: Deadline

      // Col O (index 14) = PCR only — check if already filled
      var colOVal = '';
      try {
        colOVal = sheet.getRange(i + 1, 15).getFormula() || String(row[14] || '');
      } catch(e) {}

      results.push({
        rowIndex:      i + 1,  // 1-based sheet row
        trackingCode:  String(row[0]  || ''),
        brand:         String(row[3]  || ''),
        campaignName:  String(row[4]  || ''),
        stratName:     String(row[6]  || ''),
        requesterName: String(row[7]  || ''),
        status:        String(row[9]  || ''),
        deadline:      dl instanceof Date
                         ? Utilities.formatDate(dl, tz, 'dd/MM/yyyy')
                         : (dl ? String(dl) : 'N/A'),
        hasPcr:        !!(colOVal && colOVal.trim())
      });
    }

    return results;

  } catch(e) {
    Logger.log('getKlrSubmissions error: ' + e.toString());
    return [];
  }
}

/**
 * Handles a KLR update from Campaign team:
 *  1. Creates a new Google Doc for this update (title includes update number + brand + date).
 *  2. Stores the update doc link in the next available sheet slot: P (col 16), Q (17), or R (18).
 *  3. If PCR is provided AND Col O is currently empty → also writes PCR to Col O.
 *  4. Fires a Campaign+Strategy Chat notification mentioning the strategist.
 *
 * payload: {
 *   rowIndex, requesterName, stratName, brand, campaignName, trackingCode,
 *   contextHtml,    // rich text HTML from editor (with ##IMG_N## markers)
 *   contextImages,  // [{ data, mimeType, index }] extracted images
 *   pcrType,        // 'none' | 'link' | 'file'
 *   pcrLink,        // if pcrType=link
 *   pcrAttachments, // [{ data, mimeType, name }] if pcrType=file
 *   attachments     // [{ data, mimeType, name }] supporting files
 * }
 */
function updateKlrContext(payload) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetName);
    if (!sheet) return { success: false, error: 'Sheet not found' };

    var rowIndex  = payload.rowIndex;
    var tz        = Session.getScriptTimeZone();
    var ts        = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
    var klrFolder = null;

    // ── Resolve KLR Drive folder from the Google Doc in Col N ──
    try {
      var nFormula = sheet.getRange(rowIndex, 14).getFormula();
      var nUrl     = extractHyperlinkUrl_(nFormula);
      if (nUrl) {
        var nIdMatch = nUrl.match(/\/d\/([^\/\?#]+)/);
        if (nIdMatch) {
          var nDocFile = DriveApp.getFileById(nIdMatch[1]);
          var parents  = nDocFile.getParents();
          if (parents.hasNext()) klrFolder = parents.next();
        }
      }
    } catch(fe) {
      Logger.log('updateKlrContext: folder resolve failed — ' + fe.toString());
    }

    // ── Determine next available update slot: P=16, Q=17, R=18 ──
    var updateSlot = null;
    [16, 17, 18].forEach(function(col) {
      if (updateSlot !== null) return;
      try {
        var val = sheet.getRange(rowIndex, col).getFormula() ||
                  String(sheet.getRange(rowIndex, col).getValue() || '');
        if (!val.trim()) updateSlot = col;
      } catch(e) {}
    });
    if (!updateSlot) {
      Logger.log('updateKlrContext: all update slots (P/Q/R) are full for row ' + rowIndex);
      // Fall back to R (col 18) — overwrite the oldest
      updateSlot = 18;
    }

    // ── Count existing updates to label this doc ──
    var updateNum = updateSlot - 15; // P=1, Q=2, R=3

    // ── Check Col O (15) for existing PCR ──
    var colOEmpty = false;
    try {
      var colOVal = sheet.getRange(rowIndex, 15).getFormula() ||
                    String(sheet.getRange(rowIndex, 15).getValue() || '');
      colOEmpty = !colOVal.trim();
    } catch(e) { colOEmpty = true; }

    // ── Upload PCR files/link if provided ──
    var pcrForColO      = '';   // formula to write to Col O if empty
    var pcrDocSectionHtml = ''; // HTML for the update doc's PCR section

    if (payload.pcrType === 'link' && payload.pcrLink) {
      pcrDocSectionHtml = '<h2>PCR</h2><p><a href="' + payload.pcrLink + '">' + payload.pcrLink + '</a></p>';
      if (colOEmpty) {
        pcrForColO = '=HYPERLINK("' + payload.pcrLink + '", "📂 Open PCR")';
        try { handleGoogleDocLink_(payload.pcrLink, (payload.brand || '') + ' - PCR', klrFolder || null); } catch(e) {}
      }
    } else if (payload.pcrType === 'file' && payload.pcrAttachments && payload.pcrAttachments.length > 0) {
      var pcrFileLinks = [];
      payload.pcrAttachments.forEach(function(file, idx) {
        if (!klrFolder) return;
        try {
          var blob     = Utilities.newBlob(Utilities.base64Decode(file.data), file.mimeType || 'application/octet-stream', file.name);
          var uploaded = klrFolder.createFile(blob);
          pcrFileLinks.push({ name: file.name, url: uploaded.getUrl() });
          // First file → Col O if empty
          if (idx === 0 && colOEmpty) {
            pcrForColO = '=HYPERLINK("' + uploaded.getUrl() + '", "📂 Open PCR")';
          }
        } catch(e) { Logger.log('updateKlrContext: PCR file upload failed — ' + e.toString()); }
      });
      pcrDocSectionHtml = '<h2>PCR</h2><ul>' +
        pcrFileLinks.map(function(f) { return '<li><a href="' + f.url + '">' + f.name + '</a></li>'; }).join('') +
        '</ul>';
    }

    // ── Upload supporting files ──
    var suppLinks = [];
    if (klrFolder && payload.attachments && payload.attachments.length > 0) {
      payload.attachments.forEach(function(file) {
        try {
          var blob     = Utilities.newBlob(Utilities.base64Decode(file.data), file.mimeType || 'application/octet-stream', file.name);
          var uploaded = klrFolder.createFile(blob);
          suppLinks.push({ name: file.name, url: uploaded.getUrl() });
        } catch(e) { Logger.log('updateKlrContext: file upload failed — ' + e.toString()); }
      });
    }

    // ── Extract images from contextHtml (backend fallback) ──
    var updateImages = (payload.contextImages && payload.contextImages.length > 0)
      ? payload.contextImages : [];
    var contextHtml  = payload.contextHtml || '';

    if (updateImages.length === 0 && contextHtml) {
      var imgCtr_ = 0;
      contextHtml = contextHtml.replace(
        /<img[^>]+src="(data:([^;]+);base64,([^"]+))"[^>]*\/?>/gi,
        function(match, src, mimeType, b64) {
          imgCtr_++;
          updateImages.push({ data: b64, mimeType: mimeType || 'image/png', index: imgCtr_ });
          return '<span>##IMG_' + imgCtr_ + '##</span>';
        }
      );
    }

    // ── Build update doc HTML ──
    var suppSection = suppLinks.length > 0
      ? '<h2>Supporting Files</h2><ul>' +
          suppLinks.map(function(f) { return '<li><a href="' + f.url + '">' + f.name + '</a></li>'; }).join('') +
          '</ul>'
      : '';

    var updateHtml =
      '<h1>KLR Update #' + updateNum + ' — ' + (payload.brand || '') + '</h1>' +
      '<p><b>From:</b> ' + (payload.requesterName || 'Campaign Team') + ' &nbsp;|&nbsp; <b>Date:</b> ' + ts + '</p>' +
      (contextHtml ? '<h2>Additional Context</h2>' + contextHtml : '') +
      pcrDocSectionHtml +
      suppSection;

    // ── Create update Google Doc via Drive API ──
    var updateDocUrl = '';
    var updateDocId  = '';
    if (klrFolder) {
      try {
        var updateDocTitle    = 'KLR Update #' + updateNum + ' — ' + (payload.brand || '') + ' — ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
        var updateDocBlob     = Utilities.newBlob(updateHtml, 'text/html', updateDocTitle);
        var updateDocResource = {
          title:    updateDocTitle,
          mimeType: 'application/vnd.google-apps.document',
          parents:  [{ id: klrFolder.getId() }]
        };
        var updateDocFile = Drive.Files.insert(updateDocResource, updateDocBlob, { convert: true });
        updateDocId  = updateDocFile.id;
        updateDocUrl = 'https://docs.google.com/document/d/' + updateDocId + '/edit';
        Logger.log('KLR update doc created: ' + updateDocUrl);

        // ── Insert images inline into the update doc ──
        if (updateImages.length > 0) {
          try {
            var uDoc  = DocumentApp.openById(updateDocId);
            var uBody = uDoc.getBody();
            updateImages.forEach(function(img) {
              var marker  = '##IMG_' + img.index + '##';
              var found   = uBody.findText(marker);
              var imgBlob = Utilities.newBlob(
                Utilities.base64Decode(img.data),
                img.mimeType || 'image/png',
                'image_' + img.index
              );
              if (found) {
                var para = found.getElement().getParent();
                var idx  = uBody.getChildIndex(para);
                uBody.insertImage(idx, imgBlob);
                uBody.removeChild(para);
              } else {
                uBody.appendImage(imgBlob);
              }
            });
            uDoc.saveAndClose();
            Logger.log('Inserted ' + updateImages.length + ' image(s) into update doc.');
          } catch(ie) {
            Logger.log('KLR update doc image insertion failed (non-fatal): ' + ie.toString());
          }
        }
      } catch(de) {
        Logger.log('updateKlrContext: update doc creation failed — ' + de.toString());
      }
    }

    // ── Write update doc link to next available slot (P/Q/R) ──
    if (updateDocUrl) {
      try {
        var updateLabel = 'Update #' + updateNum;
        sheet.getRange(rowIndex, updateSlot)
             .setFormula('=HYPERLINK("' + updateDocUrl + '","📝 ' + updateLabel + '")');
        Logger.log('Update doc link written to col ' + updateSlot + ' for row ' + rowIndex);
      } catch(se) {
        Logger.log('updateKlrContext: sheet write for update doc failed — ' + se.toString());
      }
    }

    // ── Write PCR to Col O if it was empty and PCR was provided ──
    if (pcrForColO && colOEmpty) {
      try {
        sheet.getRange(rowIndex, 15).setFormula(pcrForColO);
        Logger.log('PCR written to Col O for row ' + rowIndex);
      } catch(se) {
        Logger.log('updateKlrContext: Col O PCR write failed — ' + se.toString());
      }
    }

    // ── Fire Campaign+Strategy Chat notification ──
    var hasPcrUpdate = !!(pcrForColO || (payload.pcrType !== 'none' && payload.pcrType));
    var msg =
      '📎 *KLR Update #' + updateNum + ' — Additional Context*\n\n' +
      '*Brand:* '    + (payload.brand        || 'N/A') + '\n' +
      '*Campaign:* ' + (payload.campaignName || 'N/A') + '\n' +
      '*From:* '     + (payload.requesterName || 'N/A') + '\n' +
      '*Strat:* '    + getMention(payload.stratName) + ' — new update added.';

    if (contextHtml) {
      var preview = contextHtml.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().substring(0, 280);
      if (preview) msg += '\n\n*Context:*\n' + preview + (preview.length >= 280 ? '…' : '');
    }
    if (hasPcrUpdate) {
      msg += '\n📊 *PCR ' + (colOEmpty ? 'added to sheet' : 'included in update doc') + '*';
    }
    if (suppLinks.length > 0) {
      msg += '\n📎 *Files:* ' + suppLinks.map(function(f) { return f.name; }).join(', ');
    }
    if (updateDocUrl) {
      msg += '\n\n🔗 ' + updateDocUrl;
    }

    sendCampaignWebhook_(msg);

    // ── Flag the update so Strategist Dashboard shows "New Update" on KLR Brief button ──
    try {
      PropertiesService.getScriptProperties()
        .setProperty('KLR_UPDATE_' + rowIndex, new Date().toISOString());
    } catch(pe) {
      Logger.log('updateKlrContext: could not set update flag — ' + pe.toString());
    }

    return { success: true };

  } catch(e) {
    Logger.log('updateKlrContext error: ' + e.toString());
    return { success: false, error: e.message };
  }
}

/**
 * Returns the Google Chat DM conversation URL between the logged-in strategist
 * and the KLR requester identified by requesterChatId (Gaia / Chat user ID).
 *
 * Uses the Chat REST API `spaces.findDirectMessage` to look up the existing DM
 * space and returns its browser URL.  Falls back to the generic chat home URL
 * on any error so the button always does something useful.
 *
 * SETUP: In Apps Script → Project Settings → Show "appsscript.json" in editor,
 * add this scope if the API returns a 403:
 *   "https://www.googleapis.com/auth/chat.spaces"
 */
function getChatDmUrl(requesterChatId) {
  if (!requesterChatId) {
    return { success: false, url: 'https://mail.google.com/mail/u/0/#chat', error: 'No requester ID' };
  }
  try {
    var token   = ScriptApp.getOAuthToken();
    var headers = { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' };

    // 1. Try to find an existing DM space
    var findResp = UrlFetchApp.fetch(
      'https://chat.googleapis.com/v1/spaces:findDirectMessage?name=users%2F' + requesterChatId,
      { method: 'GET', headers: headers, muteHttpExceptions: true }
    );
    if (findResp.getResponseCode() === 200) {
      var spaceId = (JSON.parse(findResp.getContentText()).name || '').replace('spaces/', '');
      if (spaceId) return { success: true, url: 'https://mail.google.com/mail/u/0/#chat/dm/' + spaceId };
    }

    // 2. No existing DM — create/setup one via spaces:setup
    var setupResp = UrlFetchApp.fetch(
      'https://chat.googleapis.com/v1/spaces:setup',
      {
        method: 'POST',
        headers: headers,
        payload: JSON.stringify({
          space: { spaceType: 'DIRECT_MESSAGE' },
          memberships: [{ member: { name: 'users/' + requesterChatId, type: 'HUMAN' } }]
        }),
        muteHttpExceptions: true
      }
    );
    var setupCode = setupResp.getResponseCode();
    Logger.log('getChatDmUrl spaces:setup HTTP ' + setupCode + ' — ' + setupResp.getContentText());
    if (setupCode === 200 || setupCode === 201) {
      var spaceId = (JSON.parse(setupResp.getContentText()).name || '').replace('spaces/', '');
      if (spaceId) return { success: true, url: 'https://mail.google.com/mail/u/0/#chat/dm/' + spaceId };
    }

    return { success: false, url: 'https://mail.google.com/mail/u/0/#chat', error: 'setup API returned ' + setupCode };

  } catch(e) {
    Logger.log('getChatDmUrl error: ' + e.toString());
    return { success: false, url: 'https://mail.google.com/mail/u/0/#chat', error: e.message };
  }
}

/**
 * Clears the "new update" flag for a KLR brief.
 * Called when the strategist opens the KLR Brief from the dashboard.
 */
function clearKlrUpdate(rowIndex) {
  try {
    PropertiesService.getScriptProperties().deleteProperty('KLR_UPDATE_' + rowIndex);
    return { success: true };
  } catch(e) {
    Logger.log('clearKlrUpdate error: ' + e.toString());
    return { success: false };
  }
}

// ── Taxonomy Discovery ────────────────────────────────────────────────────

/**
 * RUN THIS ONCE FROM THE APPS SCRIPT EDITOR (not as a web app).
 * Editor → select function "discoverTaxonomy" → Run.
 *
 * Reads all Done rows from the main sheet that have both a brief link and
 * a proposal link. Extracts text from each pair (capped at 10 slides each
 * for speed), sends everything to Gemini, and asks it to discover a
 * comprehensive tag taxonomy with no cap on the number of tags.
 *
 * Result is saved as JSON to a hidden "REV Taxonomy" tab.
 * Review the Logs and that tab before running the indexing step.
 */
function discoverTaxonomy() {
  var START_MS   = new Date().getTime();
  var LIMIT_MS   = 4 * 60 * 1000; // 4-minute safety cutoff
  var MAX_PAIRS  = 100;

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetName) || ss.getSheets()[0];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('No data in sheet.'); return; }

  var numRows  = lastRow - 1;
  var values   = sheet.getRange(2, 1, numRows, 19).getValues();
  var formulas = sheet.getRange(2, 13, numRows, 6).getFormulas();

  // ── 1. Find complete pairs (Done + brief link + proposal link) ──
  var pairs = [];
  for (var i = 0; i < numRows && pairs.length < MAX_PAIRS; i++) {
    var row          = values[i];
    var status       = String(row[CONFIG.colIndices.status       - 1] || '').trim().toLowerCase();
    var proposalLink = String(row[CONFIG.colIndices.proposalLink - 1] || '').trim();

    if (status !== 'done')                           continue;
    if (!proposalLink.match(/^https?:\/\//))         continue;

    // First available brief link (cols N–R = formula offsets 1–5)
    var briefUrl = null;
    for (var li = 1; li <= 5; li++) {
      var u = extractHyperlinkUrl_(formulas[i][li]);
      if (u) { briefUrl = u; break; }
    }
    if (!briefUrl) continue;

    pairs.push({
      brand:       String(row[CONFIG.colIndices.brand   - 1] || '').trim(),
      campaign:    String(row[CONFIG.colIndices.project - 1] || '').trim(),
      industry:    String(row[8]  || '').trim(),   // col I
      deck:        String(row[1]  || '').trim(),   // col B
      briefUrl:    briefUrl,
      proposalUrl: proposalLink
    });
  }

  Logger.log('Found ' + pairs.length + ' complete brief+proposal pairs');
  if (pairs.length === 0) {
    Logger.log('Nothing to process. Check that Done rows have both brief and proposal links.');
    return;
  }

  // ── 2. Extract text from each pair ──
  var extracted = [];
  for (var j = 0; j < pairs.length; j++) {
    if (new Date().getTime() - START_MS > LIMIT_MS) {
      Logger.log('⏱ Time limit hit — processed ' + j + '/' + pairs.length + ' pairs. Running Gemini on what we have.');
      break;
    }

    var p = pairs[j];
    Logger.log('[' + (j + 1) + '/' + pairs.length + '] ' + p.brand + ' — ' + p.campaign);

    var briefText    = extractTextFromUrl_(p.briefUrl,    10);
    var proposalText = extractTextFromUrl_(p.proposalUrl, 10);
    if (!briefText && !proposalText) continue;

    extracted.push({
      brand:        p.brand,
      campaign:     p.campaign,
      industry:     p.industry,
      deck:         p.deck,
      briefText:    briefText.substring(0, 1200),
      proposalText: proposalText.substring(0, 1200)
    });
  }

  Logger.log('Successfully extracted ' + extracted.length + ' pairs');
  if (extracted.length === 0) {
    Logger.log('Extraction failed for all pairs. Check Drive permissions.');
    return;
  }

  // ── 3. Build Gemini prompt ──
  var pairsBlock = extracted.map(function(p, i) {
    return '=== PAIR ' + (i + 1) + ' ===\n' +
      'Brand: ' + p.brand + ' | Campaign: ' + p.campaign +
      ' | Industry: ' + p.industry + ' | Deck type: ' + p.deck + '\n' +
      'BRIEF:\n'    + p.briefText    + '\n' +
      'PROPOSAL:\n' + p.proposalText;
  }).join('\n\n');

  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var prompt =
    'You are analyzing ' + extracted.length + ' real brief→proposal pairs from REV Media Group, ' +
    'a Malaysian media agency. These are actual client briefs and the proposals pitched in response.\n\n' +
    'Your task: discover a COMPREHENSIVE taxonomy of tags that captures every meaningful campaign ' +
    'dimension present in this data.\n\n' +
    'RULES:\n' +
    '- Do NOT cap the number of tags. Extract every distinct, meaningful tag you find.\n' +
    '- Aim for 80–120+ unique tags across as many dimensions as the data supports.\n' +
    '- Be granular and specific — "TikTok Hashtag Challenge" is better than just "TikTok".\n' +
    '- Include Malaysian market context: local occasions, platforms, audience segments, brands.\n' +
    '- If the same concept appears under different names across pairs, unify into one clean tag.\n' +
    '- Invent new dimensions if the data warrants it — do not force everything into pre-set boxes.\n\n' +
    'Return ONLY valid JSON, no markdown fences:\n' +
    '{\n' +
    '  "dimensions": [\n' +
    '    { "name": "Campaign Occasion", "description": "What seasonal or strategic moment the campaign is built around", "tags": ["Ramadan", "Hari Raya", "CNY", ...] },\n' +
    '    ...\n' +
    '  ],\n' +
    '  "total_tags": 95,\n' +
    '  "pairs_analysed": ' + extracted.length + '\n' +
    '}\n\n' +
    'BRIEF→PROPOSAL PAIRS:\n\n' + pairsBlock;

  // ── 4. Call Gemini ──
  var resp = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0.2,
          maxOutputTokens: 8192,
          thinkingConfig: { thinkingBudget: 2048 }   // allow reasoning for taxonomy quality
        }
      }),
      muteHttpExceptions: true
    }
  );

  var respJson = JSON.parse(resp.getContentText());
  if (respJson.error) {
    Logger.log('Gemini error: ' + respJson.error.message);
    return;
  }

  var allParts = respJson.candidates[0].content.parts || [];
  var rawText  = allParts.map(function(pt) { return pt.text || ''; }).join('').trim();
  var jsonMatch = rawText.match(/\{[\s\S]*\}/);
  if (!jsonMatch) { Logger.log('No JSON found in Gemini response:\n' + rawText); return; }
  var taxonomyJson = jsonMatch[0];

  // ── 5. Save to hidden "REV Taxonomy" tab ──
  var tabName  = 'REV Taxonomy';
  var taxSheet = ss.getSheetByName(tabName);
  if (!taxSheet) { taxSheet = ss.insertSheet(tabName); taxSheet.hideSheet(); }
  taxSheet.clearContents();
  taxSheet.getRange(1, 1).setValue(taxonomyJson);
  taxSheet.getRange(2, 1).setValue(
    'Generated: ' + new Date().toISOString() +
    ' | Pairs analysed: ' + extracted.length
  );

  // ── 6. Log summary ──
  try {
    var parsed = JSON.parse(taxonomyJson);
    Logger.log('\n✅ Taxonomy saved to "' + tabName + '" tab');
    Logger.log('Total tags: ' + parsed.total_tags);
    Logger.log('Dimensions (' + parsed.dimensions.length + '):');
    parsed.dimensions.forEach(function(d) {
      Logger.log('  [' + d.tags.length + '] ' + d.name + ': ' + d.tags.join(', '));
    });
  } catch(e) {
    Logger.log('Taxonomy saved but could not parse for summary: ' + e.toString());
    Logger.log(taxonomyJson);
  }
}

/**
 * RUN FROM THE EDITOR after discoverTaxonomy() completes.
 * Reads the saved taxonomy from the "REV Taxonomy" tab and logs it
 * dimension-by-dimension in a readable format.
 */
function logTaxonomy() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('REV Taxonomy');
  if (!sheet) { Logger.log('REV Taxonomy tab not found — run discoverTaxonomy() first.'); return; }

  var raw  = sheet.getRange(1, 1).getValue();
  var meta = sheet.getRange(2, 1).getValue();
  Logger.log(meta);

  try {
    var t = JSON.parse(raw);
    Logger.log('Total tags: ' + t.total_tags + ' across ' + t.dimensions.length + ' dimensions\n');
    t.dimensions.forEach(function(d) {
      Logger.log('── ' + d.name + ' (' + d.tags.length + ' tags)');
      Logger.log('   ' + d.tags.join(', '));
    });
  } catch(e) {
    // Fallback: log in chunks if JSON is too large to parse in one go
    Logger.log('Logging in chunks (raw JSON too large to parse cleanly):');
    for (var i = 0; i < raw.length; i += 6000) {
      Logger.log(raw.substring(i, i + 6000));
    }
  }
}

/**
 * RUN FROM THE EDITOR to overwrite the REV Taxonomy tab with the
 * fully restructured taxonomy — all 14 dimensions, all original tags,
 * now organised into subgroups. Auto-computes total tag count.
 */
function saveUpdatedTaxonomy() {
  var taxonomy = {
    pairs_analysed: 94,
    dimensions: [

      // ── 1. Campaign Occasion ──────────────────────────────────────────
      {
        name: 'Campaign Occasion',
        description: 'What seasonal or strategic moment the campaign is built around',
        subgroups: [
          { name: 'Festive & Cultural',  tags: ['Ramadan & Hari Raya', 'Chinese New Year', 'Deepavali', 'National Day (Merdeka)', 'Holiday Season'] },
          { name: 'Calendar Moments',    tags: ["Parents' Day", "Mother's Day", "Father's Day", 'FIFA World Cup Season', 'World Brain Day'] },
          { name: 'Quarterly Tentpoles', tags: ['Q1 Seasonal Campaigns', 'Q2 Seasonal Campaigns', 'Q3 Seasonal Campaigns', 'Q4 Seasonal Campaigns'] },
          { name: 'Campaign Type',       tags: ['New Product Launch', 'Brand Relaunch/Restage', 'Limited Time Offer (LTO)', 'Always-On Content', 'Sustenance Content', 'Event Coverage', 'Anniversary Celebration'] }
        ]
      },

      // ── 2. Industry ───────────────────────────────────────────────────
      {
        name: 'Industry',
        description: "The client's industry sector",
        subgroups: [
          { name: 'FMCG',                     tags: ['FMCG - Food', 'FMCG - Drinks', 'FMCG - Pharmaceutical', 'FMCG - Household'] },
          { name: 'Food & Dining',             tags: ['F&B (Food & Beverage)'] },
          { name: 'Consumer & Lifestyle',      tags: ['Retail', 'Beauty/Skincare', 'Jewellery', 'E-Commerce'] },
          { name: 'Technology & Telco',        tags: ['Electronics/Gadgets', 'Telco'] },
          { name: 'Financial Services',        tags: ['Banking/Finance'] },
          { name: 'Property & Infrastructure', tags: ['Property', 'Home Improvement', 'Utilities'] },
          { name: 'Travel & Experience',       tags: ['Travel & Tourism', 'Airlines', 'Themepark/Recreational'] },
          { name: 'Regulated Industries',      tags: ['Tobacco', 'Government/Public Sector', 'Automobile'] }
        ]
      },

      // ── 3. Campaign Objective ─────────────────────────────────────────
      {
        name: 'Campaign Objective',
        description: 'What the client aims to achieve with the campaign',
        subgroups: [
          { name: 'Awareness & Reach',       tags: ['Increase Brand Awareness', 'Create Talkability/Buzz', 'Generate Virality', 'Increase Share of Voice', 'Achieve Top-of-Mind Status', 'Expand Market Reach', 'Maximize Viewership'] },
          { name: 'Brand Perception',        tags: ['Strengthen Brand Relevance', 'Build Brand Affinity', 'Improve Brand Perception', 'Reposition Brand', 'Humanize Brand', 'Reinforce Safety/Credibility', 'Strengthen Emotional Connection'] },
          { name: 'Consideration & Education', tags: ['Drive Product Consideration', 'Educate Consumers', 'Highlight Product USPs', 'Drive Trial'] },
          { name: 'Conversion & Sales',      tags: ['Drive Sales/Conversion', 'Drive Footfall (Offline)', 'Generate Leads', 'Drive App Downloads/Usage', 'Clear Product Stock'] },
          { name: 'Loyalty & Retention',     tags: ['Build Brand Loyalty', 'Build Repeat Purchase Habit', 'Increase Engagement', 'Enhance Digital Presence'] },
          { name: 'Purpose & CSR',           tags: ['Promote CSR Initiatives', 'Promote Sustainable Living', 'Promote Health Behaviors', 'Support Grassroots Initiatives', 'Financial Empowerment', 'Institutional Engagement'] },
          { name: 'Events & Activations',    tags: ['Drive Event Participation'] }
        ]
      },

      // ── 4. Target Audience ────────────────────────────────────────────
      {
        name: 'Target Audience',
        description: 'The demographic and psychographic segments the campaign targets',
        subgroups: [
          { name: 'Age & Generation',    tags: ['Gen Z (18-24)', 'Young Millennials (25-34)', 'Older Millennials (35-40)', 'Gen Y', 'Adults (18-49)', 'Adults (25-55)', 'Students'] },
          { name: 'Life Stage',          tags: ['Parents with Kids (PwK)', 'Young Families', 'Working Adults', 'Family-Oriented'] },
          { name: 'Income Tier',         tags: ['Affluent Consumers (T20)', 'Upper M40 Segment', 'Budget-Conscious Consumers'] },
          { name: 'Ethnicity & Religion', tags: ['Mass Market (All Races)', 'Malay Audience', 'Chinese Audience', 'Muslim Travelers'] },
          { name: 'Interest & Lifestyle', tags: ['Food Enthusiasts', 'Spicy Food Lovers', 'K-Culture Fans', 'Football Fans/Sports Lovers', 'Culture Enthusiasts', 'Adventure/Thrill Seekers', 'Travel Enthusiasts', 'Lifestyle-Driven Consumers', 'Tech-Savvy Individuals', 'Social Media Heavy Users', 'Health & Wellness Conscious', 'Skincare-Conscious', 'Home Cooks', 'Snackers', 'Design-Conscious Consumers', 'Value-Conscious Diners', 'Convenience-Led Consumers', 'Urban Dwellers', 'Semi-Urban Dwellers'] },
          { name: 'Professional & B2B',  tags: ['Business Owners', 'Facility Managers', 'Policy Makers/Government Sectors'] },
          { name: 'Purchase Intent',     tags: ['First Home Buyers', 'Property Investors', 'Property Upgraders', 'Nicotine Users (18+)', 'Community-Driven Individuals'] }
        ]
      },

      // ── 5. Platform Type (Digital) ────────────────────────────────────
      {
        name: 'Platform Type (Digital)',
        description: 'Specific digital channels used for campaign distribution',
        subgroups: [
          { name: 'Social Media',              tags: ['Facebook', 'Instagram', 'TikTok', 'YouTube', 'X (Twitter)', 'Threads', 'XiaoHongShu (XHS)', 'WhatsApp', 'Social Media (General)'] },
          { name: 'REV Publisher Solutions',   tags: ['Website (Publisher)', 'Lumi News App', 'Audience+ Solutions (CPC)', 'Content Retargeting'] },
          { name: 'Brand-Owned Channels',      tags: ['Microsite', 'Mobile App'] },
          { name: 'Direct Communication',      tags: ['Email Marketing (EDM)', 'Push Notifications'] },
          { name: 'Audio',                     tags: ['Podcast'] },
          { name: 'Programmatic & Display',    tags: ['Digital Display Ads'] }
        ]
      },

      // ── 6. Platform Type (Traditional/OOH) ───────────────────────────
      {
        name: 'Platform Type (Traditional/OOH)',
        description: 'Traditional media or Out-of-Home channels',
        tags: ['Television (TV3, Buletin Utama)', 'Radio (Hot FM, Fly FM)', 'Out-of-Home (OOH)', 'On-Ground Activation/Event']
      },

      // ── 7. Content Format ─────────────────────────────────────────────
      {
        name: 'Content Format',
        description: 'The type of creative asset produced',
        subgroups: [
          { name: 'Standard Video',     tags: ['Video Content (General)', 'Short-Form Video (SFV)', 'Mid-Form Video', 'Long-Form Video/Film', 'Brand Film', 'Animated Content', 'Documentary-Style Video'] },
          { name: 'Specialist Video',   tags: ['Event Coverage Video', 'Product Review Video', 'Cooking Format Video', 'Vox Pop Video', 'Teaser Content', 'Compilation Content'] },
          { name: 'Editorial',          tags: ['Editorial Article', 'Spotlight Article', 'Immersive Article', 'Listicle'] },
          { name: 'Social & Interactive', tags: ['Social Postings', 'Facebook Gallery', 'User-Generated Content (UGC)', 'Infographics'] },
          { name: 'Audio & Live',       tags: ['Podcast Series', 'Live Performances'] }
        ]
      },

      // ── 8. Creative Approach / Theme ──────────────────────────────────
      {
        name: 'Creative Approach/Theme',
        description: 'The overarching creative strategy or narrative style',
        subgroups: [
          { name: 'Narrative & Storytelling',   tags: ['Emotional Storytelling', 'Relatable Malaysian Storytelling', 'Authentic Content', 'Heritage & Origin Story', 'Transformation Story', 'Problem/Solution Narrative', 'Behind-the-Scenes Content', 'Growth Story', 'Human-Centred Design', 'Expert Endorsement'] },
          { name: 'Tone & Style',               tags: ['Humor/Comedy', 'Mystery/Intrigue Building', 'Edutainment', 'Value-for-Money Narrative', 'Lifestyle-Driven Content'] },
          { name: 'Values & Purpose',           tags: ['CSR Angle', 'Sustainability Theme', 'Circular Economy Theme', 'Modern Patriotism', 'Wellness/Self-Care Theme', 'Prevention Theme', 'Science Meets Sunnah'] },
          { name: 'Occasion & Audience Driven', tags: ['Festive-Led Ideas', 'Community-Driven Narrative', 'Family Bonding Theme', 'Gifting Theme', 'MZ-Focused Creative'] },
          { name: 'Interactive & Format-Led',   tags: ['AI-Powered Creative', 'Social-First Campaign', 'Challenge-Based Content', 'Spice Challenges', 'Street Interview', 'Interactive Elements', '"Hot Ones" Format', '"Greeters Guild" Concept'] },
          { name: 'Sensory & Experiential',     tags: ['Art & Design Integration', 'Music-Led Experience', 'Technology Focus', 'Craftsmanship'] }
        ]
      },

      // ── 9. Campaign Mechanic / Activity ──────────────────────────────
      {
        name: 'Campaign Mechanic/Activity',
        description: 'Specific interactive or promotional activities',
        subgroups: [
          { name: 'Digital Activations',  tags: ['TikTok Hashtag Challenge', 'Video Gallery', 'Digital Treasure Hunt', 'AI Contest', 'Interactive Quiz', 'Countdown Activation', 'Lead Generation Form', 'User-Generated Content (UGC) Contest', 'Pledge Campaign'] },
          { name: 'On-Ground & Events',   tags: ['On-Ground Activation', 'Pop-Up Event', 'PR Cruise Event', 'Roadshow', 'Station Activation Booths', 'Drone Show'] },
          { name: 'Promotional',          tags: ['Contest/Giveaway', 'Product Sampling', 'Blind Taste Test', 'Sponsorship Integration'] }
        ]
      },

      // ── 10. KOL / Influencer Type ─────────────────────────────────────
      {
        name: 'KOL/Influencer Type',
        description: 'Categories of Key Opinion Leaders or Content Creators',
        subgroups: [
          { name: 'By Reach/Tier',    tags: ['Celebrities', 'Macro KOLs/Influencers', 'Micro KOLs/Influencers', 'Nano KOLs/Influencers', 'KOCs (Key Opinion Consumers)'] },
          { name: 'By Content Niche', tags: ['Car Reviewers', 'Foodies/Food Reviewers', 'Beauty Influencers', 'Lifestyle Influencers', 'Singer/Songwriter KOL', 'Host/Presenter', 'Local Talent'] },
          { name: 'Emerging',         tags: ['AI KOLs/Influencers'] }
        ]
      },

      // ── 11. Malaysian Context ─────────────────────────────────────────
      {
        name: 'Malaysian Context',
        description: 'Specific local elements, cultural nuances, or market specifics',
        subgroups: [
          { name: 'Cultural Identity',       tags: ['Balik Kampung Theme', 'Citarasa Malaysia (Taste of Malaysia)', 'Malaysian Identity', 'Malaysian Unity', 'Malaysian Culture/Inspirations', 'Local Occasions (General)'] },
          { name: 'Food Culture',            tags: ['Mamaks (Local Eateries)', 'Malaysian Cuisine', 'Local Culinary Heritage'] },
          { name: 'Geographic Focus',        tags: ['Klang Valley Focus', 'Johor Focus', 'Shah Alam Focus', 'Kota Bharu Focus', 'Seremban 2 Focus', 'MVV 2045 (Malaysia Vision Valley)'] },
          { name: 'Religious & Compliance',  tags: ['Halal Certified', 'Shariah-Compliant'] },
          { name: 'Market Trends',           tags: ['Micro-Cation Trend', 'EV Market in Malaysia', 'Urban Resort Concept'] },
          { name: 'Economic Context',        tags: ['Ringgit (Currency)', 'M40/T20 Income Groups'] },
          { name: 'Local Events',            tags: ['Keretapi Sarong Event', 'Grassroots Football'] },
          { name: 'Partnerships & Initiatives', tags: ['Local Brands/Startups', 'Local Talent/Partnerships', 'Government Initiatives (SARA)', 'Public Health Emergency (Local Context)'] }
        ]
      },

      // ── 12. Product Category ──────────────────────────────────────────
      {
        name: 'Product Category',
        description: 'The specific type of product or service being advertised',
        subgroups: [
          { name: 'Fast Food / QSR',       tags: ['Fast Food (General)', 'Bundle Meals', 'Crispy Tenders', 'Soccer Ball Hashbrown', 'Beef Burgers', 'Paratha', 'Paratha Pocket', 'Spicy Chicken Popcorn', 'Family Meals'] },
          { name: 'Beverages',             tags: ['Ginger Brew (Drink)', 'Daily Beverage (Functional)', 'Coffee (Brand/Product)', 'Canned Milk'] },
          { name: 'FMCG Snacks & Food',    tags: ['Instant Noodles', 'Coconut Biscuit', 'Snacks (General)', 'Chips/Crisps', 'Limited-Edition Packaging'] },
          { name: 'Home Appliances',       tags: ['Home Appliances (General)', 'Water Purifier', 'Air Purifier', 'Styler (Appliance)', 'Vacuum Cleaner', 'Refrigerator', 'Washing Machine', 'Dryer', 'Air Conditioner', 'Smart TV'] },
          { name: 'Technology & Smartphones', tags: ['Smartphones (General)', 'Mobile Series (S, A, Z)', 'AI Features (Tech)'] },
          { name: 'Automotive',            tags: ['Automobiles (General)', 'SUV', 'Crossover EV', 'Large MPV EV'] },
          { name: 'Property',              tags: ['Property Developments (General)', 'Residential Units', 'Condominiums/Suites', 'Township Development', 'EV-Ready Homes'] },
          { name: 'Beauty & Personal Care', tags: ['UV Serum', 'Micellar Water', 'Deodorant', 'Jewellery (Gold, Jade)'] },
          { name: 'Health & Wellness',     tags: ['Multivitamin', 'Toothpaste'] },
          { name: 'Travel & Hospitality',  tags: ['Airline Routes', 'Low-Cost Airline Services', 'Hotels/Resorts', 'Staycation Packages', 'MICE Venues', 'Travel Packages', 'Islamic Pilgrimage Services (Umrah, Hajj)'] },
          { name: 'Regulated Products',    tags: ['Unbranded Tobacco Alternatives (IQOS)', 'Water Utility Services', 'Flooring Solutions'] },
          { name: 'Financial & Digital',   tags: ['Financial Services/Retail Banking', 'E-Commerce Platform', 'Digital App (Brand-Specific)'] },
          { name: 'Recreation & Education', tags: ['Themepark Rides', 'School Challenge/Competition'] }
        ]
      },

      // ── 13. Success Metrics ───────────────────────────────────────────
      {
        name: 'Success Metrics',
        description: "How the campaign's performance is measured",
        subgroups: [
          { name: 'Awareness & Reach',         tags: ['Impressions', 'Reach', 'Brand Recall', 'Brand Sentiment', 'Talkability'] },
          { name: 'Engagement',                tags: ['Engagement Rate', 'Video Views', 'Page Views'] },
          { name: 'Performance & Conversion',  tags: ['Clicks (CTR)', 'Website Traffic', 'Conversion Rate', 'ROI (Return on Investment)'] },
          { name: 'Business Outcomes',         tags: ['Sales Uplift', 'App Downloads', 'Footfall', 'Registrations', 'Qualified Leads'] }
        ]
      },

      // ── 14. Challenges / Barriers ─────────────────────────────────────
      {
        name: 'Challenges/Barriers',
        description: 'Obstacles the campaign aims to overcome',
        subgroups: [
          { name: 'Brand & Perception',    tags: ['Low Brand Awareness', 'Declining Brand Strength', 'Brand Confusion', 'Negative Brand Perception', 'Maintaining Relevance', 'Translating Legacy to Modern', 'Perception of Declining Quality'] },
          { name: 'Market & Competitive',  tags: ['Crowded Market/Content Space', 'Difficulty Standing Out'] },
          { name: 'Conversion & Retention', tags: ['Ineffective Conversion', 'Lack of Repeat Purchase'] },
          { name: 'Product & Trust',       tags: ['Product Misconceptions', 'Product Safety/Halal Validity Concerns', 'Intimidation Barrier (Travel)'] },
          { name: 'Executional',           tags: ['Balancing Short-Term/Long-Term Goals', 'Ensuring Consistency', 'Limited Budget', 'No Existing Assets', 'Strict Advertising Regulations (Tobacco)'] },
          { name: 'Information & Education', tags: ['Information Overload/Misinformation', 'Public Health Message Fatigue'] }
        ]
      }

    ] // end dimensions
  };

  // Auto-compute total tags
  var total = 0;
  taxonomy.dimensions.forEach(function(d) {
    if (d.subgroups) {
      d.subgroups.forEach(function(sg) { total += sg.tags.length; });
    } else if (d.tags) {
      total += d.tags.length;
    }
  });
  taxonomy.total_tags = total;

  // Save to REV Taxonomy tab
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('REV Taxonomy');
  if (!sheet) { sheet = ss.insertSheet('REV Taxonomy'); sheet.hideSheet(); }
  sheet.clearContents();
  sheet.getRange(1, 1).setValue(JSON.stringify(taxonomy, null, 2));
  sheet.getRange(2, 1).setValue('Updated with subgroups: ' + new Date().toISOString());

  // Log summary
  Logger.log('✅ Taxonomy saved with subgroups');
  Logger.log('Total tags: ' + total + ' across ' + taxonomy.dimensions.length + ' dimensions\n');
  taxonomy.dimensions.forEach(function(d) {
    if (d.subgroups) {
      var n = d.subgroups.reduce(function(a, sg) { return a + sg.tags.length; }, 0);
      Logger.log('  ' + d.name + ' — ' + n + ' tags in ' + d.subgroups.length + ' subgroups');
      d.subgroups.forEach(function(sg) {
        Logger.log('    [' + sg.tags.length + '] ' + sg.name + ': ' + sg.tags.join(', '));
      });
    } else {
      Logger.log('  ' + d.name + ' — ' + (d.tags || []).length + ' tags (flat)');
      Logger.log('    ' + (d.tags || []).join(', '));
    }
  });
}

/**
 * Extracts text from a Google Slides, Docs, Sheets, or PDF/image URL.
 * Handles all brief formats used at REV Media.
 * @param {string} url       Full Google Drive/Docs URL
 * @param {number} maxSlides Max slides to read (Slides only)
 */
function extractTextFromUrl_(url, maxSlides) {
  if (!url) return '';
  try {
    var idMatch = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!idMatch) {
      console.log('🔍 extractTextFromUrl_: no file ID found in URL: ' + url.substring(0, 80));
      return '';
    }
    var fileId = idMatch[1];
    console.log('🔍 extractTextFromUrl_: fileId=' + fileId + ' url=' + url.substring(0, 80));

    // ── Google Slides ──────────────────────────────────────────────────────
    if (url.indexOf('presentation') !== -1) {
      console.log('🔍 extractTextFromUrl_: type=SLIDES');
      var pres   = SlidesApp.openById(fileId);
      var slides = pres.getSlides();
      var limit  = Math.min(slides.length, maxSlides || 10);
      console.log('🔍 extractTextFromUrl_: opened OK, slides=' + slides.length + ', reading up to ' + limit);
      var texts  = [];
      for (var i = 0; i < limit; i++) {
        var st = [];
        slides[i].getPageElements().forEach(function(el) {
          try {
            if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
              var t = el.asShape().getText().asString().trim();
              if (t) st.push(t);
            }
          } catch(e) {}
        });
        if (st.length) texts.push(st.join(' '));
      }
      var result = texts.join(' | ');
      console.log('🔍 extractTextFromUrl_: extracted ' + result.length + ' chars from SLIDES');
      return result;

    // ── Google Docs ────────────────────────────────────────────────────────
    } else if (url.indexOf('document') !== -1) {
      console.log('🔍 extractTextFromUrl_: type=DOCS');
      var docText = DocumentApp.openById(fileId).getBody().getText().substring(0, 3000);
      console.log('🔍 extractTextFromUrl_: extracted ' + docText.length + ' chars from DOCS');
      return docText;

    // ── Google Sheets ──────────────────────────────────────────────────────
    } else if (url.indexOf('spreadsheets') !== -1) {
      console.log('🔍 extractTextFromUrl_: type=SHEETS');
      var ss       = SpreadsheetApp.openById(fileId);
      var sheets   = ss.getSheets();
      var texts    = [];
      var sheetCap = Math.min(sheets.length, 3);   // first 3 tabs
      for (var s = 0; s < sheetCap; s++) {
        var data     = sheets[s].getDataRange().getValues();
        var rowTexts = [];
        var rowCap   = Math.min(data.length, 80);
        for (var r = 0; r < rowCap; r++) {
          var rowStr = data[r].filter(function(c) { return c !== '' && c !== null; }).join(' | ');
          if (rowStr.trim()) rowTexts.push(rowStr);
        }
        if (rowTexts.length) {
          texts.push('Sheet "' + sheets[s].getName() + '":\n' + rowTexts.join('\n'));
        }
      }
      var sheetResult = texts.join('\n\n').substring(0, 3000);
      console.log('🔍 extractTextFromUrl_: extracted ' + sheetResult.length + ' chars from SHEETS');
      return sheetResult;

    // ── PDF / image / Drive file by ID ────────────────────────────────────
    } else {
      var file = DriveApp.getFileById(fileId);
      var mime = file.getMimeType();
      console.log('🔍 extractTextFromUrl_: type=DRIVE FILE mimeType=' + mime);

      // Google Slides shared via /file/d/ URL
      if (mime === 'application/vnd.google-apps.presentation') {
        console.log('🔍 extractTextFromUrl_: redirecting to SlidesApp (mime match)');
        var pres2   = SlidesApp.openById(fileId);
        var slides2 = pres2.getSlides();
        var limit2  = Math.min(slides2.length, maxSlides || 10);
        var texts2  = [];
        for (var j = 0; j < limit2; j++) {
          var st2 = [];
          slides2[j].getPageElements().forEach(function(el) {
            try {
              if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
                var t2 = el.asShape().getText().asString().trim();
                if (t2) st2.push(t2);
              }
            } catch(e) {}
          });
          if (st2.length) texts2.push(st2.join(' '));
        }
        var res2 = texts2.join(' | ');
        console.log('🔍 extractTextFromUrl_: extracted ' + res2.length + ' chars via SlidesApp (mime)');
        return res2;

      // Google Docs shared via /file/d/ URL
      } else if (mime === 'application/vnd.google-apps.document') {
        console.log('🔍 extractTextFromUrl_: redirecting to DocumentApp (mime match)');
        var docText2 = DocumentApp.openById(fileId).getBody().getText().substring(0, 3000);
        console.log('🔍 extractTextFromUrl_: extracted ' + docText2.length + ' chars via DocumentApp (mime)');
        return docText2;

      // PDF / image — try OCR, fallback to DocumentApp if blob is actually a Google Doc
      } else {
        console.log('🔍 extractTextFromUrl_: OCR path for mime=' + mime);
        var blob2    = file.getBlob();
        var blobMime = blob2.getContentType();
        console.log('🔍 extractTextFromUrl_: blob contentType=' + blobMime);

        // If the blob is actually a Google Doc internally, open it directly
        if (blobMime === 'application/vnd.google-apps.document' ||
            blobMime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
          console.log('🔍 extractTextFromUrl_: blob is Doc type, using DocumentApp fallback');
          try {
            var docFallback = DocumentApp.openById(fileId).getBody().getText().substring(0, 3000);
            console.log('🔍 extractTextFromUrl_: extracted ' + docFallback.length + ' chars via DocumentApp fallback');
            return docFallback;
          } catch(docErr) {
            console.log('❌ extractTextFromUrl_: DocumentApp fallback also failed: ' + docErr.toString());
            return '';
          }
        }

        var resource = { title: '[TEMP OCR] ' + file.getName(),
                         mimeType: 'application/vnd.google-apps.document' };
        var tempFile = Drive.Files.insert(resource, blob2, { ocr: true, ocrLanguage: 'en' });
        var ocrText  = DocumentApp.openById(tempFile.id).getBody().getText();
        DriveApp.getFileById(tempFile.id).setTrashed(true);
        console.log('🔍 extractTextFromUrl_: extracted ' + ocrText.length + ' chars via OCR');
        return ocrText.substring(0, 3000);
      }
    }

  } catch(e) {
    console.log('❌ extractTextFromUrl_ FAILED url=' + url.substring(0, 80) + ' error=' + e.toString());
    Logger.log('extractTextFromUrl_ failed (' + url.substring(0, 60) + '): ' + e.toString());
  }
  return '';
}

/**
 * Extracts and combines text from multiple brief URLs.
 * Used to give Gemini full context of the current brief, including all attached files.
 * @param {string[]} briefUrls  Array of brief file URLs (up to 5)
 * @param {number}   maxSlides  Max slides per file
 * @param {number}   maxChars   Max chars per file (default 1500)
 */
function extractBriefText_(briefUrls, maxSlides, maxChars) {
  if (!briefUrls || !briefUrls.length) return '';
  maxChars = maxChars || 1500;
  var parts = [];
  briefUrls.forEach(function(url, i) {
    if (!url) return;
    var text = extractTextFromUrl_(url, maxSlides || 10);
    if (text) parts.push('--- Brief File ' + (i + 1) + ' ---\n' + text.substring(0, maxChars));
  });
  return parts.join('\n\n');
}

// ── Related Proposals ─────────────────────────────────────────────────────

/**
 * Archive tab definitions. Tab names must match exactly what exists in the sheet.
 * dataRow = first row containing data (1-based). maxCol = rightmost column needed.
 */
var ARCHIVE_TABS_ = [
  { name: 'Archive 2023',      year: '2023',    dataRow: 2,
    cols: { brand:4, campaign:5, industry:11, deck:2, status:12, link:14 }, maxCol: 14 },
  { name: 'Archive 2024',      year: '2024',    dataRow: 2,
    cols: { brand:4, campaign:5, industry:10, deck:2, status:11, link:13 }, maxCol: 13 },
  { name: 'Archive 2025/2026', year: '2025/26', dataRow: 3,
    cols: { brand:5, campaign:6, industry:11, deck:2, status:12, link:14 }, maxCol: 14 },
];

// ── Proposal Index — Tag-Based Matching ──────────────────────────────────────

var INDEX_SHEET_NAME_ = 'Proposal Index';
var BATCH_SIZE_       = 40;   // proposals per runIndexingJob() run

/**
 * Loads the REV Taxonomy JSON from the hidden tab.
 * Result is cached in CacheService for 6 hours so repeated calls are free.
 */
function loadTaxonomy_() {
  var cache  = CacheService.getScriptCache();
  var cached = cache.get('REV_TAXONOMY');
  if (cached) return JSON.parse(cached);

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('REV Taxonomy');
  if (!sheet) throw new Error('REV Taxonomy tab not found. Run saveUpdatedTaxonomy() first.');

  var raw      = sheet.getRange(1, 1).getValue();
  var taxonomy = JSON.parse(raw);
  cache.put('REV_TAXONOMY', raw, 21600);   // 6-hour cache
  return taxonomy;
}

/**
 * Builds a compact text representation of the taxonomy for Gemini prompts.
 * Format: "DimensionName > SubgroupName: tag1, tag2, ..."
 * Flat dimensions (no subgroups): "DimensionName (flat): tag1, tag2, ..."
 */
function buildTaxonomyText_() {
  var tax   = loadTaxonomy_();
  var lines = [];
  tax.dimensions.forEach(function(d) {
    if (d.subgroups) {
      d.subgroups.forEach(function(sg) {
        lines.push(d.name + ' > ' + sg.name + ': ' + sg.tags.join(', '));
      });
    } else if (d.tags) {
      lines.push(d.name + ' (flat): ' + d.tags.join(', '));
    }
  });
  return lines.join('\n');
}

/**
 * Sends proposal/brief text to Gemini and returns an array of flat tag strings.
 * Each tag has format: "DimensionName::SubgroupName::TagName"
 * or "DimensionName::TagName" for flat dimensions.
 */
function tagWithGemini_(text, brand, campaign, industry, deck) {
  var apiKey  = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY not set in Script Properties');

  var taxText = buildTaxonomyText_();

  var prompt =
    'You are tagging a media proposal/brief from a Malaysian media agency for a searchable index.\n\n' +
    'METADATA:\n' +
    'Brand: ' + brand + '\n' +
    'Campaign: ' + campaign + '\n' +
    'Industry: ' + (industry || 'N/A') + '\n' +
    'Deck Type: ' + (deck || 'N/A') + '\n\n' +
    'TEXT (extracted from slides):\n' + String(text || '').substring(0, 4000) + '\n\n' +
    'TAXONOMY (Dimension > Subgroup: available tags):\n' + taxText + '\n\n' +
    'Instructions:\n' +
    '- Assign ALL relevant tags from the taxonomy above.\n' +
    '- Be thorough — cover every dimension that clearly applies.\n' +
    '- Use the EXACT tag names as they appear in the taxonomy.\n' +
    '- For dimensions with subgroups: use format "DimensionName::SubgroupName::TagName"\n' +
    '- For flat dimensions (no subgroup): use format "DimensionName::TagName"\n\n' +
    'Return ONLY a JSON array of strings — no markdown, no explanation:\n' +
    '["Campaign Occasion::Festive & Cultural::Ramadan & Hari Raya", "Industry::FMCG::FMCG - Food", ...]';

  var resp = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 1024,
          thinkingConfig: { thinkingBudget: 0 } }
      }),
      muteHttpExceptions: true
    }
  );

  var json = JSON.parse(resp.getContentText());
  if (json.error) throw new Error('Gemini: ' + json.error.message);

  var parts = (json.candidates[0].content.parts || [])
    .map(function(p) { return p.text || ''; }).join('').trim();
  var match = parts.match(/\[[\s\S]*\]/);
  if (!match) return [];
  try { return JSON.parse(match[0]); } catch(e) { return []; }
}

/**
 * Computes a 0–100 match score between two tag arrays.
 * Scoring: exact tag = 3pts, same subgroup = 2pts, same dimension = 1pt.
 * Score = total / (briefTags.length × 3) × 100.
 */
function computeTagScore_(briefTags, proposalTags) {
  if (!briefTags || !briefTags.length || !proposalTags || !proposalTags.length) return 0;

  var propExact     = {};
  var propSubgroup  = {};
  var propDimension = {};

  proposalTags.forEach(function(t) {
    propExact[t] = true;
    var p = t.split('::');
    if (p.length >= 2) propSubgroup[p[0] + '::' + p[1]] = true;
    propDimension[p[0]] = true;
  });

  var total = 0;
  var max   = briefTags.length * 3;

  briefTags.forEach(function(t) {
    var p = t.split('::');
    if      (propExact[t])                                                 total += 3;
    else if (p.length >= 2 && propSubgroup[p[0] + '::' + (p[1] || '')])  total += 2;
    else if (propDimension[p[0]])                                          total += 1;
  });

  return max > 0 ? Math.round(total / max * 100) : 0;
}

/**
 * Returns the tag leaf names (last segment after ::) that appear in both arrays.
 * Limited to 5 so the justification sentence stays readable.
 */
function getMatchingTagNames_(briefTags, proposalTags) {
  var propSet = {};
  proposalTags.forEach(function(t) { propSet[t] = true; });
  return briefTags
    .filter(function(t) { return propSet[t]; })
    .map(function(t) { return t.split('::').pop(); })
    .slice(0, 5);
}

/**
 * Returns (or creates) the hidden "Proposal Index" sheet.
 * Columns: Source Tab | Brand | Campaign | Year | Deck | URL | Tags | Indexed At
 */
function getOrCreateIndexSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(INDEX_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(INDEX_SHEET_NAME_);
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Source Tab', 'Brand', 'Campaign', 'Year', 'Deck', 'URL', 'Tags', 'Indexed At'
    ]]);
    sheet.hideSheet();
    Logger.log('Created hidden "' + INDEX_SHEET_NAME_ + '" sheet.');
  }
  return sheet;
}

/**
 * RUN FROM EDITOR (or time trigger) — processes up to BATCH_SIZE_ archive
 * proposals per invocation, tags them with Gemini, and appends to the index.
 * Stops automatically when all proposals are indexed.
 */
function runIndexingJob() {
  // Prevent overlapping runs if a previous trigger fires while this one is still executing
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('⏳ Another indexing run is still in progress. Skipping this trigger fire.');
    return;
  }

  var START_MS  = new Date().getTime();
  var LIMIT_MS  = 4 * 60 * 1000;   // stop after 4 minutes (safe margin from 6-min limit)
  var ss        = SpreadsheetApp.getActiveSpreadsheet();

  // Load already-indexed URLs into a lookup set
  var indexSheet  = getOrCreateIndexSheet_();
  var indexedUrls = {};
  var lastIndexRow = indexSheet.getLastRow();
  if (lastIndexRow > 1) {
    indexSheet.getRange(2, 6, lastIndexRow - 1, 1).getValues()
      .forEach(function(r) { if (r[0]) indexedUrls[r[0]] = true; });
  }

  // Collect all unindexed proposals from archive tabs
  var queue = [];
  ARCHIVE_TABS_.forEach(function(tab) {
    readArchiveTab_(ss, tab).forEach(function(r) {
      if (!indexedUrls[r.link]) queue.push({
        source: tab.name, year: tab.year,
        brand: r.brand, campaign: r.campaign,
        industry: r.industry, deck: r.deck, url: r.link
      });
    });
  });

  if (queue.length === 0) {
    Logger.log('✅ All proposals indexed. Total: ' + (lastIndexRow - 1));
    stopIndexingTrigger_();
    return;
  }

  Logger.log('Queue: ' + queue.length + ' unindexed. Processing up to ' + BATCH_SIZE_ + ' this run.');

  var newRows   = [];
  var processed = 0;

  for (var i = 0; i < queue.length && processed < BATCH_SIZE_; i++) {
    if (new Date().getTime() - START_MS > LIMIT_MS) {
      Logger.log('⏱ Time limit reached. Stopping at ' + processed + '.');
      break;
    }

    var p = queue[i];
    Logger.log('[' + (processed + 1) + '/' + Math.min(queue.length, BATCH_SIZE_) + '] ' +
      p.brand + ' — ' + p.campaign.substring(0, 40));

    var text = extractTextFromUrl_(p.url, 20);
    var tags = [];
    var note = '';

    if (text) {
      try {
        tags = tagWithGemini_(text, p.brand, p.campaign, p.industry, p.deck);
        Logger.log('  → ' + tags.length + ' tags');
      } catch(e) {
        note = 'tag_error: ' + e.message.substring(0, 80);
        Logger.log('  ⚠ ' + note);
      }
    } else {
      note = 'no_text';
      Logger.log('  ⚠ Could not extract text (no access or unsupported format)');
    }

    newRows.push([
      p.source, p.brand, p.campaign, p.year, p.deck,
      p.url,
      tags.length ? JSON.stringify(tags) : (note || '[]'),
      new Date().toISOString()
    ]);
    processed++;
  }

  if (newRows.length > 0) {
    indexSheet.getRange(indexSheet.getLastRow() + 1, 1, newRows.length, 8).setValues(newRows);
    Logger.log('✅ Indexed ' + newRows.length + ' this run. Total: ' + (lastIndexRow - 1 + newRows.length));
  }

  lock.releaseLock();
}

/**
 * Creates a time-based trigger that runs runIndexingJob() every 5 minutes.
 * Run this ONCE from the editor to start background indexing.
 * It removes itself when the queue is empty.
 */
function setupIndexingTrigger() {
  stopIndexingTrigger_();
  ScriptApp.newTrigger('runIndexingJob').timeBased().everyMinutes(1).create();
  Logger.log('✅ Trigger created. runIndexingJob() will fire every 5 minutes.');
  Logger.log('   It stops automatically when all proposals are indexed.');
  Logger.log('   Or call stopIndexingTrigger() to cancel early.');
}

/** Call from the editor to cancel background indexing. */
function stopIndexingTrigger() {
  stopIndexingTrigger_();
  Logger.log('Indexing trigger removed.');
}

function stopIndexingTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runIndexingJob') ScriptApp.deleteTrigger(t);
  });
}

/**
 * Run from the editor to see indexing progress.
 * Logs: total indexed, breakdown by archive tab, and how many have empty tags.
 */
/**
 * Run from the editor to clear all Related Proposals and Insights cache entries.
 * Use this to force fresh results after updating taxonomy, insights logic, etc.
 */
function clearRelatedCache() {
  // CacheService doesn't support wildcard removal.
  // Strategy: store all wow_ / rel_ keys in a registry so we can nuke them precisely.
  // Fallback: also try common static keys.
  var cache = CacheService.getScriptCache();

  // Pull the key registry (written by cacheWithRegistry_ below)
  var registry = [];
  try { registry = JSON.parse(cache.get('CACHE_KEY_REGISTRY') || '[]'); } catch(e) {}

  // Always include static keys
  var staticKeys = ['REV_TAXONOMY', 'CACHE_KEY_REGISTRY'];
  var allKeys = staticKeys.concat(registry);

  // Remove in batches of 100 (Apps Script limit)
  while (allKeys.length > 0) {
    cache.removeAll(allKeys.splice(0, 100));
  }

  Logger.log('✅ Cache cleared (' + (staticKeys.length + registry.length) + ' keys). Next click will run fresh.');
}

/**
 * Writes a value to script cache AND registers the key so clearRelatedCache() can find it.
 */
function putCacheWithRegistry_(cache, key, value, ttl) {
  cache.put(key, value, ttl);
  var registry = [];
  try { registry = JSON.parse(cache.get('CACHE_KEY_REGISTRY') || '[]'); } catch(e) {}
  if (registry.indexOf(key) === -1) {
    registry.push(key);
    // Keep registry itself alive; trim if huge
    if (registry.length > 500) registry = registry.slice(-500);
    cache.put('CACHE_KEY_REGISTRY', JSON.stringify(registry), 21600);
  }
}

function getIndexStats() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(INDEX_SHEET_NAME_);
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log('Proposal Index is empty. Run setupIndexingTrigger() to start indexing.');
    return;
  }
  var count   = sheet.getLastRow() - 1;
  var data    = sheet.getRange(2, 1, count, 7).getValues();
  var sources = {};
  var empty   = 0;
  data.forEach(function(r) {
    var src = String(r[0] || 'Unknown');
    sources[src] = (sources[src] || 0) + 1;
    var tagStr = String(r[6] || '');
    try {
      if (JSON.parse(tagStr).length === 0) empty++;
    } catch(e) { empty++; }
  });
  Logger.log('📊 Proposal Index Stats');
  Logger.log('   Total indexed: ' + count);
  for (var s in sources) Logger.log('   ' + s + ': ' + sources[s]);
  Logger.log('   No tags (failed/inaccessible): ' + empty +
    ' (' + Math.round(empty / count * 100) + '%)');
}

// ── Auto-index on Done ───────────────────────────────────────────────────────

/**
 * Shared helper — indexes a single row from the main Work List sheet
 * if status = Done AND a valid proposal link exists AND not already indexed.
 * Called from both the onEdit trigger (sheet path) and submitProposal() (web app path).
 */
function indexRowIfReady_(sheet, rowIndex) {
  // Read enough columns to get all fields we need (proposal link is the rightmost at col K=11)
  var numCols = Math.max(CONFIG.colIndices.proposalLink || 11, 11);
  var row = sheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];

  var status       = String(row[CONFIG.colIndices.status       - 1] || '').trim().toLowerCase();
  var proposalLink = String(row[CONFIG.colIndices.proposalLink - 1] || '').trim();
  var brand        = String(row[CONFIG.colIndices.brand        - 1] || '').trim();
  var campaign     = String(row[CONFIG.colIndices.project      - 1] || '').trim();
  var industry     = String(row[8] || '').trim();   // col I
  var deck         = String(row[1] || '').trim();   // col B

  // Only proceed if Done + valid Google Slides/Docs link
  if (status !== 'done') return;
  if (!proposalLink || !proposalLink.match(/^https?:\/\//)) return;

  // Skip if already in the index
  var indexSheet = getOrCreateIndexSheet_();
  var lastRow    = indexSheet.getLastRow();
  if (lastRow > 1) {
    var urls = indexSheet.getRange(2, 6, lastRow - 1, 1).getValues();
    for (var i = 0; i < urls.length; i++) {
      if (urls[i][0] === proposalLink) {
        Logger.log('indexRowIfReady_: already indexed — ' + brand + ' / ' + campaign);
        return;
      }
    }
  }

  Logger.log('Auto-indexing: ' + brand + ' — ' + campaign);
  var text = extractTextFromUrl_(proposalLink, 20);
  var tags = [];
  if (text) {
    try {
      tags = tagWithGemini_(text, brand, campaign, industry, deck);
      Logger.log('  → ' + tags.length + ' tags');
    } catch(e) {
      Logger.log('  ⚠ Tagging failed: ' + e.toString());
    }
  } else {
    Logger.log('  ⚠ Could not extract text (unsupported format or no access)');
  }

  indexSheet.getRange(indexSheet.getLastRow() + 1, 1, 1, 8).setValues([[
    '2025/2026 Work List', brand, campaign, '2025/26', deck,
    proposalLink,
    tags.length ? JSON.stringify(tags) : '[]',
    new Date().toISOString()
  ]]);
  Logger.log('✅ Auto-indexed: ' + brand + ' — ' + campaign);
}

/**
 * Installable onEdit trigger handler.
 * Fires when a strategist directly edits the Status or Proposal Link column in the sheet.
 * Run setupOnDoneTrigger() once from the editor to install this.
 */
function onProposalDone(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.sheetName) return;   // only watch the main sheet

    var col = e.range.getColumn();
    var row = e.range.getRow();
    if (row < 2) return;  // skip header

    // Only act when Status or Proposal Link column is touched
    if (col !== CONFIG.colIndices.status && col !== CONFIG.colIndices.proposalLink) return;

    indexRowIfReady_(sheet, row);
  } catch(err) {
    Logger.log('onProposalDone error: ' + err.toString());
  }
}

/**
 * Run ONCE from the editor to install the onEdit trigger.
 * After this, any direct sheet edit to Status/Proposal Link will auto-index if Done.
 */
function setupOnDoneTrigger() {
  // Remove existing copies to avoid duplicate triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onProposalDone') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onProposalDone')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  Logger.log('✅ onProposalDone trigger installed.');
  Logger.log('   New proposals marked Done in the sheet will be indexed automatically.');
}

// ── Proposal Insights — Wow Factor ───────────────────────────────────────────

/**
 * Extracts text from the top 3 matched proposals and asks Gemini what made
 * each one stand out relative to the current brief.
 * Called from the frontend immediately after renderRelated() displays the 5 cards.
 *
 * @param {Array}  results   Top 5 result objects from getRelatedProposals()
 * @param {string} brand
 * @param {string} campaign
 * @param {string} industry
 * @returns {Array} [{index:0, wow:"..."}, {index:1, wow:"..."}, {index:2, wow:"..."}]
 */
function getProposalInsights(results, brand, campaign, industry, briefUrls) {
  try {
    if (!results || results.length === 0) return [];

    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) throw new Error('GEMINI_API_KEY not set');

    // Normalise briefUrls
    if (typeof briefUrls === 'string') briefUrls = briefUrls ? [briefUrls] : [];
    briefUrls = (briefUrls || []).filter(Boolean);

    // Cache check — key on the 3 proposal URLs so repeated clicks are instant
    var top3    = results.slice(0, 3);
    var cacheKey = 'wow2_' + top3.map(function(r) { return r.link; }).join('|')
      .replace(/[^a-z0-9|]/gi, '_').substring(0, 200);
    var cache   = CacheService.getScriptCache();
    var cached  = cache.get(cacheKey);
    if (cached) {
      console.log('⚡ getProposalInsights: CACHE HIT for key=' + cacheKey);
      return JSON.parse(cached);
    }
    console.log('🔄 getProposalInsights: CACHE MISS — fetching fresh. key=' + cacheKey);

    // Extract brief text directly here (avoids passing large strings through frontend)
    var briefText = extractBriefText_(briefUrls, 10, 1500);
    console.log('📋 getProposalInsights brief source: ' + (briefText ? 'FILE TEXT (' + briefText.length + ' chars)' : 'METADATA ONLY'));

    // Extract text from top 3 proposals (25 slides each)
    var proposalBlocks = top3.map(function(r, i) {
      console.log('🔍 getProposalInsights: extracting proposal ' + (i+1) + ' — ' + r.brand + ' / ' + r.campaign + ' link=' + (r.link || '').substring(0, 60));
      var text = extractTextFromUrl_(r.link, 25);
      console.log('🔍 getProposalInsights: proposal ' + (i+1) + ' extracted ' + (text ? text.length : 0) + ' chars');
      return '--- Proposal ' + (i + 1) + ': ' + r.brand + ' / ' + r.campaign +
        ' [' + r.year + '] ---\n' + (text ? text.substring(0, 3000) : '[could not extract text]');
    }).join('\n\n');

    var briefSection = briefText
      ? 'CURRENT BRIEF (extracted from brief files):\n' + briefText.substring(0, 3000)
      : 'CURRENT BRIEF:\nBrand: ' + brand + '\nCampaign: ' + campaign + '\nIndustry: ' + (industry || 'N/A');

    var prompt =
      'You are a Creative Director with 10 international awards across Cannes Lions and the Effie Awards. ' +
      'You think in equal parts creative bravery and business effectiveness. ' +
      'A junior strategist has come to you before writing a proposal and needs your honest, unfiltered read on 3 past decks.\n\n' +
      briefSection + '\n\n' +
      'PAST PROPOSALS (top 3 tag-matched results):\n\n' + proposalBlocks + '\n\n' +
      'For each past proposal, write exactly 2 sentences in your voice as this Creative Director:\n\n' +
      'Sentence 1 — THE SHARPEST IDEA: What is genuinely creative in this deck? ' +
      'Look for: the human truth it is built on, the mechanic that surprised people, the format that broke category norms, ' +
      'or the strategic angle that reframed the problem. Reference specific content from the deck. ' +
      'If the idea is derivative or not actually interesting, say so honestly.\n\n' +
      'Sentence 2 — STEAL OR SKIP: Is there something in this proposal that directly maps to the current brief? ' +
      'Name the exact element to adapt — the emotional territory, the platform mechanic, the audience insight, the campaign structure. ' +
      'If nothing translates cleanly, say exactly: "Nothing maps cleanly here — open it for reference only, not inspiration." ' +
      'Never force a connection that does not exist.\n\n' +
      'Rules:\n' +
      '- Think carefully before committing to an answer. Only write what you are fully confident about.\n' +
      '- Speak directly — you are talking to a strategist, not writing a report\n' +
      '- Zero filler words: no "innovative", "compelling", "impactful", "strong strategy", "leverages", "seamlessly"\n' +
      '- Every sentence must be grounded in actual content from the deck and the brief\n' +
      '- If a proposal deck could not be read, write: "Could not preview this deck — open it directly"\n\n' +
      'Return ONLY a JSON array, no markdown, no extra text:\n' +
      '[{"index":0,"wow":"..."},{"index":1,"wow":"..."},{"index":2,"wow":"..."}]';

    var resp = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 1, maxOutputTokens: 2048,
            thinkingConfig: { thinkingBudget: 2048 } }
        }),
        muteHttpExceptions: true
      }
    );

    var json = JSON.parse(resp.getContentText());
    if (json.error) throw new Error('Gemini: ' + json.error.message);

    var parts = (json.candidates[0].content.parts || [])
      .map(function(p) { return p.text || ''; }).join('').trim();
    var match = parts.match(/\[[\s\S]*\]/);
    if (!match) {
      console.log('❌ getProposalInsights: no JSON array found in response. raw=' + parts.substring(0, 300));
      return [];
    }

    var jsonStr = match[0];
    var insights;
    try {
      insights = JSON.parse(jsonStr);
    } catch(parseErr) {
      console.log('⚠️ getProposalInsights: JSON parse failed (' + parseErr.toString() + '), attempting cleanup...');
      // Walk the string char-by-char, only escaping control chars that appear inside string values
      try {
        var cleaned = '';
        var inStr   = false;
        var esc     = false;
        for (var ci = 0; ci < jsonStr.length; ci++) {
          var ch = jsonStr[ci];
          if (esc) { cleaned += ch; esc = false; continue; }
          if (ch === '\\' && inStr) { cleaned += ch; esc = true; continue; }
          if (ch === '"') { cleaned += ch; inStr = !inStr; continue; }
          if (inStr) {
            var code = ch.charCodeAt(0);
            if (ch === '\n') { cleaned += '\\n'; continue; }
            if (ch === '\r') { cleaned += '\\r'; continue; }
            if (ch === '\t') { cleaned += '\\t'; continue; }
            if (code < 0x20 || code === 0x7F) continue; // drop other control chars
          }
          cleaned += ch;
        }
        insights = JSON.parse(cleaned);
        console.log('✅ getProposalInsights: cleanup parse succeeded');
      } catch(parseErr2) {
        console.log('❌ getProposalInsights: cleanup parse also failed. raw=' + jsonStr.substring(0, 300));
        Logger.log('getProposalInsights JSON parse failed: ' + parseErr2.toString());
        return [];
      }
    }

    putCacheWithRegistry_(cache, cacheKey, JSON.stringify(insights), 21600);
    return insights;

  } catch(e) {
    Logger.log('getProposalInsights error: ' + e.toString());
    return [];
  }
}

/**
 * Analyses a single proposal on demand (for results #4 and #5).
 * Called when the strategist clicks "✦ Analyse this proposal" on card 4 or 5.
 *
 * @param {Object} result  Single result object {brand, campaign, year, link, ...}
 * @param {string} brand
 * @param {string} campaign
 * @param {string} industry
 * @returns {string} The wow factor sentence(s)
 */
function getSingleInsight(result, brand, campaign, industry, briefUrls) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) throw new Error('GEMINI_API_KEY not set');

    // Normalise briefUrls
    if (typeof briefUrls === 'string') briefUrls = briefUrls ? [briefUrls] : [];
    briefUrls = (briefUrls || []).filter(Boolean);

    var cacheKey = 'wow3s_' + (result.link || '').replace(/[^a-z0-9]/gi, '_').substring(0, 150);
    var cache    = CacheService.getScriptCache();
    var cached   = cache.get(cacheKey);
    if (cached) return cached;

    var text = extractTextFromUrl_(result.link, 25);

    // Extract brief text directly (avoids passing large strings through frontend)
    var briefText = extractBriefText_(briefUrls, 10, 1500);

    var briefSection = briefText
      ? 'CURRENT BRIEF (extracted from brief files):\n' + briefText.substring(0, 2000)
      : 'CURRENT BRIEF:\nBrand: ' + brand + ', Campaign: ' + campaign + ', Industry: ' + (industry || 'N/A');

    var prompt =
      'You are a Creative Director with 10 international awards across Cannes Lions and the Effie Awards. ' +
      'You think in equal parts creative bravery and business effectiveness. ' +
      'A junior strategist has come to you before writing a proposal and needs your honest read on one past deck.\n\n' +
      briefSection + '\n\n' +
      'PAST PROPOSAL:\n' +
      result.brand + ' / ' + result.campaign + ' [' + result.year + ']\n\n' +
      (text ? text.substring(0, 3500) : '[could not extract text]') + '\n\n' +
      'Write exactly 2 sentences in your voice as this Creative Director:\n\n' +
      'Sentence 1 — THE SHARPEST IDEA: What is genuinely creative in this deck? ' +
      'The human truth it is built on, the mechanic that surprised, the format that broke category norms. ' +
      'Reference specific content. If it is not actually interesting, say so.\n\n' +
      'Sentence 2 — STEAL OR SKIP: Name the exact element to adapt to the current brief — ' +
      'the emotional territory, the mechanic, the insight, the structure. ' +
      'If nothing maps cleanly, say: "Nothing maps cleanly here — open it for reference only, not inspiration." ' +
      'Never force a connection.\n\n' +
      'Rules: think carefully, only commit when confident, zero filler words, ground every sentence in real content.\n\n' +
      'Return ONLY the 2 plain text sentences, no JSON, no markdown, no labels.';

    var resp = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 1, maxOutputTokens: 2048,
            thinkingConfig: { thinkingBudget: 2048 } }
        }),
        muteHttpExceptions: true
      }
    );

    var json = JSON.parse(resp.getContentText());
    if (json.error) throw new Error('Gemini: ' + json.error.message);

    var wow = (json.candidates[0].content.parts || [])
      .filter(function(p) { return p.text && !p.thought; })
      .map(function(p) { return p.text; }).join('').trim();

    putCacheWithRegistry_(cache, cacheKey, wow, 21600);
    return wow;

  } catch(e) {
    Logger.log('getSingleInsight error: ' + e.toString());
    return 'Could not analyse this proposal.';
  }
}

// ── Related Proposals — main entry point ─────────────────────────────────────

/**
 * Called from the Strategist Dashboard.
 * If the Proposal Index is populated, uses tag-based matching (more accurate).
 * Falls back to the legacy metadata + Gemini re-ranking approach if index is empty.
 *
 * @param {string}   brand
 * @param {string}   campaign
 * @param {string}   industry
 * @param {string}   deckType
 * @param {string[]} briefUrls  All brief file URLs — Slides, Docs, Sheets, PDFs all handled.
 *                              Text is extracted from all files and combined for Gemini context.
 */
function getRelatedProposals(brand, campaign, industry, deckType, briefUrls) {
  try {
    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var indexSheet = ss.getSheetByName(INDEX_SHEET_NAME_);
    var indexCount = indexSheet ? indexSheet.getLastRow() - 1 : 0;

    // Normalise briefUrls — accept string (legacy) or array
    if (typeof briefUrls === 'string') briefUrls = briefUrls ? [briefUrls] : [];
    briefUrls = (briefUrls || []).filter(Boolean);

    // Fall back to legacy approach if index not yet built
    if (indexCount < 1) {
      Logger.log('Index empty — using legacy approach. Run setupIndexingTrigger() to build index.');
      return getRelatedProposalsLegacy_(brand, campaign, industry, deckType);
    }

    // Cache check
    var cacheKey = 'relv3_' + (brand + '|' + industry + '|' + briefUrls.join(','))
      .toLowerCase().replace(/[^a-z0-9|,]/g, '_').substring(0, 200);
    var cache    = CacheService.getScriptCache();

    console.log('📎 briefUrls received (' + briefUrls.length + '): ' + JSON.stringify(briefUrls));

    var cached   = cache.get(cacheKey);
    if (cached) {
      console.log('⚡ Returning cached result for: ' + cacheKey.substring(0, 60));
      return JSON.parse(cached);
    }

    // Extract combined text from ALL brief files (Slides, Docs, Sheets, PDFs)
    var briefText = extractBriefText_(briefUrls, 10, 1500);
    console.log('📄 briefText extracted: ' + (briefText ? briefText.length + ' chars' : 'EMPTY — no brief files or extraction failed'));
    var tagInput  =
      (briefText ? briefText + '\n\n' : '') +
      'Brand: ' + brand + '\nCampaign: ' + campaign +
      '\nIndustry: ' + (industry || '') + '\nDeck: ' + (deckType || '');

    var briefTags = tagWithGemini_(tagInput, brand, campaign, industry, deckType);
    if (!briefTags.length) return getRelatedProposalsLegacy_(brand, campaign, industry, deckType);

    // Load index and score every row
    var rows   = indexSheet.getRange(2, 1, indexCount, 8).getValues();
    var scored = [];

    rows.forEach(function(row) {
      var url     = String(row[5] || '');
      var tagsRaw = String(row[6] || '[]');
      if (!url || tagsRaw.indexOf('[') === -1) return;

      var proposalTags = [];
      try { proposalTags = JSON.parse(tagsRaw); } catch(e) {}
      if (!proposalTags.length) return;

      var score = computeTagScore_(briefTags, proposalTags);
      if (score === 0) return;

      scored.push({
        brand:     String(row[1] || ''),
        campaign:  String(row[2] || ''),
        year:      String(row[3] || ''),
        deck:      String(row[4] || ''),
        link:      url,
        score:     score,
        matchTags: getMatchingTagNames_(briefTags, proposalTags)
      });
    });

    scored.sort(function(a, b) { return b.score - a.score; });

    var results = scored.slice(0, 5).map(function(r) {
      var tagLine = r.matchTags.length
        ? r.matchTags.join(' · ')
        : 'similar campaign profile';
      return {
        brand:    r.brand,
        campaign: r.campaign,
        year:     r.year,
        deck:     r.deck,
        link:     r.link,
        score:    r.score,
        reason:   r.score + '% match — ' + tagLine
      };
    });

    // Return briefText so insights calls can reuse it without re-extracting
    var response = { success: true, results: results, briefText: briefText };
    putCacheWithRegistry_(cache, cacheKey, JSON.stringify(response), 21600);
    return response;

  } catch(e) {
    Logger.log('getRelatedProposals error: ' + e.toString());
    return { success: false, error: e.message, results: [] };
  }
}

/**
 * Legacy fallback: metadata pre-filter + Gemini re-ranking.
 * Used when the Proposal Index hasn't been built yet.
 */
function getRelatedProposalsLegacy_(brand, campaign, industry, deckType) {
  try {
    var cacheKey = 'rel_' + (brand + '|' + industry)
      .toLowerCase().replace(/[^a-z0-9|]/g, '_').substring(0, 200);
    var cache  = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    // 2. Pre-filter candidates from all archive tabs
    var ss           = SpreadsheetApp.getActiveSpreadsheet();
    var candidates   = [];
    var normIndustry = normalizeIndustry_(industry);

    ARCHIVE_TABS_.forEach(function(tab) {
      readArchiveTab_(ss, tab).forEach(function(r) {
        var brandMatch    = r.brand.toLowerCase() === brand.toLowerCase();
        var industryMatch = normIndustry !== '' &&
                            normalizeIndustry_(r.industry) === normIndustry;
        if (brandMatch || industryMatch) {
          r.brandMatch = brandMatch;
          candidates.push(r);
        }
      });
    });

    if (candidates.length === 0) {
      var empty = { success: true, results: [] };
      putCacheWithRegistry_(cache, cacheKey, JSON.stringify(empty), 21600);
      return empty;
    }

    // 3. Brand matches first; cap at 60 for Gemini
    candidates.sort(function(a, b) {
      return (b.brandMatch ? 1 : 0) - (a.brandMatch ? 1 : 0);
    });
    var subset = candidates.slice(0, 60);

    // 4. Gemini re-ranking
    var results  = rankWithGemini_(brand, campaign, industry, deckType, subset);
    var response = { success: true, results: results };
    putCacheWithRegistry_(cache, cacheKey, JSON.stringify(response), 21600);
    return response;

  } catch(e) {
    Logger.log('getRelatedProposalsLegacy_ error: ' + e.toString());
    return { success: false, error: e.message, results: [] };
  }
}

/**
 * Reads one archive tab and returns an array of normalised row objects.
 * Only includes rows that are status="done" AND have a valid proposal URL.
 */
function readArchiveTab_(ss, tab) {
  var sheet = ss.getSheetByName(tab.name);
  if (!sheet) { Logger.log('Archive tab not found: ' + tab.name); return []; }

  var lastRow = sheet.getLastRow();
  if (lastRow < tab.dataRow) return [];

  var numRows  = lastRow - tab.dataRow + 1;
  var values   = sheet.getRange(tab.dataRow, 1, numRows, tab.maxCol).getValues();
  var linkFmls = sheet.getRange(tab.dataRow, tab.cols.link, numRows, 1).getFormulas();

  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row      = values[i];
    var brand    = String(row[tab.cols.brand    - 1] || '').trim();
    var campaign = String(row[tab.cols.campaign - 1] || '').trim();
    var industry = String(row[tab.cols.industry - 1] || '').trim();
    var deck     = String(row[tab.cols.deck     - 1] || '').trim();
    var status   = String(row[tab.cols.status   - 1] || '').trim().toLowerCase();
    var linkVal  = String(row[tab.cols.link     - 1] || '').trim();
    var link     = extractHyperlinkUrl_(linkFmls[i][0]) ||
                   (linkVal.match(/^https?:\/\//) ? linkVal : '');

    if (!brand || !campaign || !link || status !== 'done') continue;
    out.push({ brand:brand, campaign:campaign, industry:industry,
               deck:deck, link:link, year:tab.year });
  }
  return out;
}

/**
 * Maps messy/inconsistent industry labels (60+ variants across years)
 * to a normalised key for cross-year matching.
 */
function normalizeIndustry_(raw) {
  if (!raw) return '';
  var s = raw.toLowerCase().trim();
  if (/f\s*&\s*b|food.*bev|bev.*food/.test(s))        return 'fnb';
  if (/fmcg.*food|food.*fmcg/.test(s))                return 'fmcg_food';
  if (/fmcg.*drink|drink.*fmcg/.test(s))              return 'fmcg_drinks';
  if (/fmcg.*household|household.*fmcg/.test(s))      return 'fmcg_household';
  if (/fmcg.*pharma|pharma.*fmcg/.test(s))            return 'fmcg_pharma';
  if (/\bfmcg\b/.test(s))                             return 'fmcg';
  if (/beauty|skincare|cosmetic/.test(s))              return 'beauty';
  if (/tech|gadget|electron/.test(s))                  return 'tech';
  if (/bank|financ|insurance/.test(s))                 return 'banking';
  if (/property|real.?estate/.test(s))                 return 'property';
  if (/telco|telecom/.test(s))                         return 'telco';
  if (/\bgov\b|glc|government/.test(s))               return 'gov';
  if (/travel|tourism|airline|hotel/.test(s))          return 'travel';
  if (/auto|car|vehicle/.test(s))                      return 'auto';
  if (/retail|fashion|apparel/.test(s))                return 'retail';
  if (/health|wellness/.test(s))                       return 'health';
  if (/education|school|university/.test(s))           return 'education';
  return s.replace(/[^a-z0-9]/g, '');
}

/**
 * Calls Gemini 2.0 Flash to re-rank up to 60 candidate proposals
 * and returns the top 5 most relevant ones with a one-line reason each.
 */
function rankWithGemini_(brand, campaign, industry, deckType, candidates) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY not set in Script Properties');

  var list = candidates.map(function(c, i) {
    return (i + 1) + '. [' + c.year + '] ' + c.brand + ' — ' + c.campaign +
      (c.deck ? ' (' + c.deck + ')' : '');
  }).join('\n');

  var prompt =
    'You are helping a media strategist find relevant past proposals to reference.\n\n' +
    'Current brief:\n' +
    'Brand: ' + brand + '\n' +
    'Campaign: ' + campaign + '\n' +
    'Industry: ' + industry + '\n' +
    'Deck type: ' + (deckType || 'N/A') + '\n\n' +
    'From the list below, pick the 5 most useful past proposals to reference ' +
    '(or fewer if less than 5 exist). Prioritise: same brand, same seasonal ' +
    'occasion (Ramadan, CNY, etc.), similar campaign mechanic, same industry, ' +
    'or similar deck type.\n\n' +
    list + '\n\n' +
    'Reply ONLY with a valid JSON array — no markdown fences, no extra text:\n' +
    '[{"index":1,"reason":"one short sentence explaining why relevant"}]\n' +
    'Indexes are 1-based numbers from the list above.';

  var resp = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0.1,
          maxOutputTokens: 512,
          thinkingConfig: { thinkingBudget: 0 }  // disable thinking — not needed for ranking
        }
      }),
      muteHttpExceptions: true
    }
  );

  var json = JSON.parse(resp.getContentText());
  if (json.error) throw new Error('Gemini API error: ' + json.error.message);

  // Concatenate all parts (thinking models may split output across parts)
  var allParts = json.candidates[0].content.parts || [];
  var text = allParts.map(function(p) { return p.text || ''; }).join('').trim();

  // Extract the JSON array — handles preamble, postamble, and markdown fences
  var arrayMatch = text.match(/\[[\s\S]*\]/);
  if (!arrayMatch) throw new Error('Gemini returned no JSON array. Raw: ' + text.substring(0, 200));
  var ranked = JSON.parse(arrayMatch[0]);

  return ranked.slice(0, 5).map(function(r) {
    var c = candidates[r.index - 1];
    if (!c) return null;
    return {
      brand:    c.brand,
      campaign: c.campaign,
      year:     c.year,
      deck:     c.deck,
      link:     c.link,
      reason:   r.reason || ''
    };
  }).filter(Boolean);
}

// -----------------------------------------------------------------------

/**
 * Converts Quill's HTML output to clean plain text.
 * Preserves paragraph breaks and list items.
 */
function htmlToPlainText_(html) {
  if (!html) return '';
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/h[1-6]>/gi, '\n\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<li>/gi, '• ')
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// ============================================================
// PRODUCT INDEX — scan product deck folders, tag + describe
// ============================================================

var PRODUCT_FOLDERS_ = [
  { folderId: '1v8PRNp5_-mNoDBm_nExBINKTR7Im53Yk', year: '2026' },
  { folderId: '1yL7qaLrUIyfwRMmJYJWNxhYy8lcJx7-S', year: '2025' },
  { folderId: '1v7kPsdf_y7fDYvAKykYT6KdxQ5TGC3Wx', year: '2024' }
];

var PRODUCT_INDEX_SHEET_ = 'Product Index';
var PRODUCT_BATCH_SIZE_  = 15;

/**
 * Returns (or creates) the hidden Product Index sheet.
 * Columns: Year | Name | Link | Tags (JSON) | Description | Indexed At
 */
function getOrCreateProductIndexSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PRODUCT_INDEX_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(PRODUCT_INDEX_SHEET_);
    sheet.hideSheet();
    sheet.getRange(1, 1, 1, 7).setValues([[
      'Year', 'Name', 'Link', 'Tags', 'Description', 'Indexed At', 'Concept'
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Scans the 3 product folders (recursively, including subfolders) and
 * returns all presentation files (Google Slides + uploaded PPTX) not
 * yet indexed (checked against the Product Index sheet).
 */
function getUnindexedProductFiles_() {
  var sheet   = getOrCreateProductIndexSheet_();
  var lastRow = sheet.getLastRow();
  var indexed = {};
  if (lastRow > 1) {
    sheet.getRange(2, 3, lastRow - 1, 1).getValues().forEach(function(r) {
      if (r[0]) indexed[String(r[0]).trim()] = true;
    });
  }

  // Collect all presentation files from a folder and its subfolders recursively
  function scanFolder(folder, year, out) {
    // Native Google Slides
    var iter = folder.getFilesByType(MimeType.GOOGLE_SLIDES);
    while (iter.hasNext()) {
      var file = iter.next();
      var url  = 'https://docs.google.com/presentation/d/' + file.getId() + '/edit';
      if (!indexed[url]) {
        out.push({ name: file.getName(), url: url, id: file.getId(), year: year });
      }
    }
    // Uploaded PowerPoint (.pptx)
    var pptIter = folder.getFilesByType('application/vnd.openxmlformats-officedocument.presentationml.presentation');
    while (pptIter.hasNext()) {
      var pptFile = pptIter.next();
      var pptUrl  = 'https://drive.google.com/file/d/' + pptFile.getId() + '/view';
      if (!indexed[pptUrl]) {
        out.push({ name: pptFile.getName(), url: pptUrl, id: pptFile.getId(), year: year });
      }
    }
    // Recurse into subfolders
    var subIter = folder.getFolders();
    while (subIter.hasNext()) {
      scanFolder(subIter.next(), year, out);
    }
  }

  var files = [];
  PRODUCT_FOLDERS_.forEach(function(f) {
    try {
      var folder = DriveApp.getFolderById(f.folderId);
      scanFolder(folder, f.year, files);
    } catch(e) {
      Logger.log('⚠️ Could not scan folder ' + f.folderId + ': ' + e.toString());
    }
  });
  return files;
}

/**
 * One Gemini call that returns tags, a 1-sentence description, and a concept label.
 * Returns { tags: [...], description: "...", concept: "..." }
 */
function tagProductWithGemini_(text, name, year) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY not set');

  var taxonomy = buildTaxonomyText_();

  var prompt =
    'You are analysing an internal media/content product deck from a Malaysian media company.\n\n' +
    'PRODUCT NAME: ' + name + ' (' + year + ')\n\n' +
    'DECK CONTENT (first ' + text.length + ' chars):\n' + text + '\n\n' +
    'TAXONOMY:\n' + taxonomy + '\n\n' +
    'Task 1 — TAGS: Select the most relevant tags from the taxonomy above that describe what this product is and who it is best suited for. ' +
    'Use the exact format "Dimension::Subgroup::Tag". Return 5–15 tags.\n\n' +
    'Task 2 — DESCRIPTION: Write exactly 1 sentence (max 20 words) describing what this product is and what it achieves for a brand. ' +
    'Be specific — name the key mechanic or format.\n\n' +
    'Task 3 — CONCEPT: Write a 3-6 word Title Case label for the PRIMARY use case this product is best pitched for. ' +
    'Think of it as the campaign goal a brand would have, not the product name. ' +
    'Examples: "Awareness for Product Launch", "Social Commerce Drive", "Brand Building with Creators", "Festive Season Engagement", "SME Digital Presence". ' +
    'Do NOT repeat the product name. Do NOT use generic labels like "Media Package".\n\n' +
    'Return ONLY valid JSON, no markdown:\n' +
    '{"tags":["Dimension::Subgroup::Tag",...],"description":"One sentence here.","concept":"Title Case Label Here"}';

  var resp = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.2, maxOutputTokens: 512,
          thinkingConfig: { thinkingBudget: 0 } }
      }),
      muteHttpExceptions: true
    }
  );

  var json = JSON.parse(resp.getContentText());
  if (json.error) throw new Error('Gemini: ' + json.error.message);

  var raw = (json.candidates[0].content.parts || [])
    .filter(function(p) { return p.text && !p.thought; })
    .map(function(p) { return p.text; }).join('').trim();

  var match = raw.match(/\{[\s\S]*\}/);
  if (!match) throw new Error('No JSON in response: ' + raw.substring(0, 100));
  return JSON.parse(match[0]);
}

/**
 * Main batch indexing job for product decks.
 * Run via trigger or manually. Processes PRODUCT_BATCH_SIZE_ files per run.
 */
function runProductIndexingJob() {
  // Use a timestamped Properties flag so we don't clash with runIndexingJob.
  // Treat as stale if older than 5 minutes (guards against killed runs that skip finally).
  var props     = PropertiesService.getScriptProperties();
  var flagTime  = parseInt(props.getProperty('PRODUCT_INDEX_RUNNING') || '0', 10);
  var stale     = (Date.now() - flagTime) > 5 * 60 * 1000;
  if (flagTime && !stale) {
    Logger.log('⏳ Product indexing already running. Skipping.');
    return;
  }
  props.setProperty('PRODUCT_INDEX_RUNNING', String(Date.now()));

  try {
    var files = getUnindexedProductFiles_();
    if (files.length === 0) {
      Logger.log('✅ All product decks are indexed.');
      stopProductIndexingTrigger_();
      return;
    }

    Logger.log('📦 ' + files.length + ' product files unindexed. Processing up to ' + PRODUCT_BATCH_SIZE_ + '...');

    var sheet = getOrCreateProductIndexSheet_();
    var batch = files.slice(0, PRODUCT_BATCH_SIZE_);
    var rows  = [];

    batch.forEach(function(f) {
      try {
        var text   = extractTextFromUrl_(f.url, 15) || '';
        var input  = text.substring(0, 2000);
        var result = tagProductWithGemini_(input, f.name, f.year);
        rows.push([
          f.year,
          f.name,
          f.url,
          JSON.stringify(result.tags || []),
          result.description || '',
          new Date().toISOString(),
          result.concept || ''
        ]);
        Logger.log('✅ Indexed: ' + f.name + ' (' + (result.tags || []).length + ' tags) — ' + (result.concept || ''));
      } catch(e) {
        Logger.log('❌ Failed: ' + f.name + ' — ' + e.toString());
        rows.push([f.year, f.name, f.url, '[]', '', new Date().toISOString(), '']);
      }
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    }

    var remaining = files.length - batch.length;
    Logger.log('📊 Done this run. ~' + remaining + ' product files remaining.');

  } finally {
    PropertiesService.getScriptProperties().deleteProperty('PRODUCT_INDEX_RUNNING');
  }
}

/**
 * Sets up a 5-minute trigger for runProductIndexingJob.
 * Run once manually.
 */
function setupProductIndexingTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runProductIndexingJob') {
      Logger.log('Trigger already exists.');
      return;
    }
  }
  ScriptApp.newTrigger('runProductIndexingJob').timeBased().everyMinutes(5).create();
  Logger.log('✅ Product indexing trigger created (every 5 min).');
}

function stopProductIndexingTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runProductIndexingJob') ScriptApp.deleteTrigger(t);
  });
}

function stopProductIndexingTrigger() {
  stopProductIndexingTrigger_();
  Logger.log('🛑 Product indexing trigger removed.');
}

/**
 * Clears the Product Index sheet (keeps header) and resets the running flag.
 * Run this before re-indexing from scratch (e.g. after fixing the scanner).
 */
function resetProductIndex() {
  var sheet   = getOrCreateProductIndexSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  PropertiesService.getScriptProperties().deleteProperty('PRODUCT_INDEX_RUNNING');
  Logger.log('🗑️ Product Index cleared. Ready to re-index.');
}

/**
 * Returns top product deck matches for a given brief.
 * Same tag-matching logic as getRelatedProposals — scores each product
 * in the Product Index against the brief's discovered tags.
 * Returns array of { name, url, year, description, matchTags, score }.
 */
function getRelevantProducts(brand, campaign, industry, briefUrls) {
  try {
    if (typeof briefUrls === 'string') briefUrls = briefUrls ? [briefUrls] : [];
    briefUrls = (briefUrls || []).filter(Boolean);

    var cacheKey = 'prod1_' + (brand + '|' + industry + '|' + briefUrls.join(','))
      .toLowerCase().replace(/[^a-z0-9|,]/g, '_').substring(0, 200);
    var cache  = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch(e) {}
    }

    var sheet   = getOrCreateProductIndexSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    // Tag the brief (reuse same pipeline as getRelatedProposals)
    var briefText = extractBriefText_(briefUrls, 10, 1500);
    var tagInput  =
      (briefText ? briefText + '\n\n' : '') +
      'Brand: ' + brand + '\nCampaign: ' + campaign +
      '\nIndustry: ' + (industry || '');
    var briefTags = tagWithGemini_(tagInput, brand, campaign, industry, '');
    if (!briefTags.length) return [];

    // Score every product in the index
    var rows   = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    var scored = [];
    rows.forEach(function(row) {
      var year = String(row[0] || '');
      var name = String(row[1] || '');
      var url  = String(row[2] || '');
      var desc = String(row[4] || '');
      var concept = String(row[6] || '');
      var productTags = [];
      try { productTags = JSON.parse(String(row[3] || '[]')); } catch(e) {}
      if (!productTags.length || !url) return;

      var score     = computeTagScore_(briefTags, productTags);
      if (score === 0) return;
      var matchTags = getMatchingTagNames_(briefTags, productTags);
      scored.push({ name: name, url: url, year: year, description: desc,
                    concept: concept, matchTags: matchTags, score: score });
    });

    scored.sort(function(a, b) { return b.score - a.score; });
    var top = scored.slice(0, 6);

    putCacheWithRegistry_(cache, cacheKey, JSON.stringify(top), 1800);
    return top;
  } catch(e) {
    Logger.log('getRelevantProducts error: ' + e.toString());
    return [];
  }
}

/**
 * Returns all indexed products for the ambient marquee.
 * Lightweight — no Gemini call, just reads the sheet.
 */
function getAllProducts() {
  try {
    var sheet   = getOrCreateProductIndexSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var rows    = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    var result  = [];
    rows.forEach(function(row) {
      var url  = String(row[2] || '');
      var name = String(row[1] || '');
      if (!url || !name) return;
      result.push({
        year:        String(row[0] || ''),
        name:        name,
        url:         url,
        description: String(row[4] || ''),
        concept:     String(row[6] || '')
      });
    });
    return result;
  } catch(e) {
    Logger.log('getAllProducts error: ' + e.toString());
    return [];
  }
}

/**
 * Logs stats on the Product Index sheet.
 */
function getProductIndexStats() {
  var sheet = getOrCreateProductIndexSheet_();
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log('Product Index is empty. Run setupProductIndexingTrigger() to start.');
    return;
  }
  var count  = sheet.getLastRow() - 1;
  var data   = sheet.getRange(2, 1, count, 5).getValues();
  var byYear = {};
  var noDesc = 0;
  data.forEach(function(r) {
    var yr = String(r[0] || 'Unknown');
    byYear[yr] = (byYear[yr] || 0) + 1;
    if (!r[4]) noDesc++;
  });
  Logger.log('📊 Product Index: ' + count + ' total');
  for (var y in byYear) Logger.log('   ' + y + ': ' + byYear[y]);
  Logger.log('   Missing description: ' + noDesc);
}


// ════════════════════════════════════════════════════════════════════════════
// CASE STUDY & KEY LEARNING INDEX
// ════════════════════════════════════════════════════════════════════════════
/**
 * ── CONFIG ──────────────────────────────────────────────────────────────────
 * Paste the Google Drive folder IDs for each year below.
 * To get a folder ID: open the folder in Drive → copy the ID from the URL
 * (the long string after /folders/)
 */
var CS_FOLDERS_ = [
  { year: '2023',    folderId: '17A_ca8nUXmLOIzCAwoqaUQOnCrQ8P0vA' },
  { year: '2024',    folderId: '1OlhlZzIqLr3GKyB8hT8zuoaMN26VsVSF' },
  { year: '2025',    folderId: '1rWHXdSGs1y9My5VLpNkqvJTwPAv9pmZe' },
  { year: '2026',    folderId: '1ikfykQjv17mC63CfodLeX_r6kb6QbDAw' },
];

var CS_INDEX_SHEET_NAME_ = 'CS & KL Index';
var CS_BATCH_SIZE_       = 10;   // files per run (lower than proposals — Gemini summary adds time)

// ── Sheet helper ─────────────────────────────────────────────────────────────

function getOrCreateCaseStudyIndexSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CS_INDEX_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(CS_INDEX_SHEET_NAME_);
    sheet.getRange(1, 1, 1, 9).setValues([[
      'Year', 'File Name', 'URL', 'Type', 'Brand', 'Campaign', 'Tags', 'Summary', 'Indexed At'
    ]]);
    sheet.hideSheet();
    Logger.log('Created hidden "' + CS_INDEX_SHEET_NAME_ + '" sheet.');
  }
  return sheet;
}

// ── Gemini: classify + tag + summarise in one call ───────────────────────────

function classifyAndSummariseForCsIndex_(text, fileName) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY not set');

  var taxonomy = loadTaxonomy_();
  var taxStr   = JSON.stringify(taxonomy).substring(0, 4000);

  var prompt =
    'You are indexing a strategic document from a media agency archive.\n\n' +
    'FILE NAME: ' + fileName + '\n\n' +
    'DOCUMENT TEXT (first ~2000 chars):\n' + text.substring(0, 2000) + '\n\n' +
    'Return a JSON object with exactly these fields:\n' +
    '{\n' +
    '  "type": "Case Study" OR "Key Learning",\n' +
    '  "brand": "client or brand name",\n' +
    '  "campaign": "campaign name",\n' +
    '  "tags": [...],\n' +
    '  "summary": "2 plain sentences."\n' +
    '}\n\n' +
    'Rules for type:\n' +
    '- "Case Study": document is about a campaign that ran, performed well, and shows proof of results or creative success.\n' +
    '- "Key Learning": document is primarily about a campaign that underperformed, faced issues, or exists to document lessons and how to avoid repeating mistakes.\n' +
    '- When in doubt, use the file name and opening paragraphs to decide.\n\n' +
    'Rules for tags: use the taxonomy below. Return an array of tag strings in the format "Dimension::Subgroup::Tag".\n' +
    'Taxonomy: ' + taxStr + '\n\n' +
    'Rules for summary — write exactly 2 sentences. Voice and focus differ by type:\n\n' +
    'If type = "Case Study":\n' +
    '  Persona: You are a senior Campaign Strategist at a Malaysian media agency with 15 years experience. ' +
    'You read case studies not to celebrate wins but to extract what is genuinely replicable.\n' +
    '  Sentence 1 — THE IDEA: Name the single most creative or strategic idea in this campaign. ' +
    'Be specific — name the mechanic, format, or human insight it was built on. ' +
    'Banned phrases: "leveraged social media", "engaged audiences", "drove awareness", "innovative approach".\n' +
    '  Sentence 2 — THE PROOF: State one concrete result with a number if available (reach, views, CTR, sales lift, % above target). ' +
    'If no number exists, name the one thing that specifically worked and why it worked in the Malaysian market context.\n\n' +
    'If type = "Key Learning":\n' +
    '  Persona: You are a tough but fair Campaign Director reviewing a post-mortem in a Malaysian media agency. ' +
    'You cut through the justifications to name exactly what went wrong — no softening, no corporate language.\n' +
    '  Sentence 1 — WHAT FAILED: Name specifically what underperformed — the platform, format, audience assumption, or execution gap — ' +
    'with a metric if available. Do not write vague statements like "results were below expectations".\n' +
    '  Sentence 2 — THE FIX: Give one actionable recommendation specific enough that a strategist can apply it to their next brief ' +
    'without opening this deck. Start with a verb: "Avoid...", "Replace...", "Always test...", "Cap spend on..." etc.\n\n' +
    'Universal rules for both types:\n' +
    '- Maximum 40 words total across both sentences\n' +
    '- Ground every word in actual content from the document — no assumptions\n' +
    '- Zero filler: no "compelling", "impactful", "strong performance", "key takeaway"\n\n' +
    'Return ONLY the JSON object, no markdown, no code fences.';

  var resp = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.3, maxOutputTokens: 1024,
          thinkingConfig: { thinkingBudget: 512 } }
      }),
      muteHttpExceptions: true
    }
  );

  var json = JSON.parse(resp.getContentText());
  if (json.error) throw new Error('Gemini: ' + json.error.message);

  var raw = (json.candidates[0].content.parts || [])
    .filter(function(p) { return p.text && !p.thought; })
    .map(function(p) { return p.text; }).join('').trim();

  // Strip markdown fences if present
  raw = raw.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/i, '').trim();

  // Robust parse (same char-walker as getProposalInsights)
  var parsed;
  try {
    parsed = JSON.parse(raw);
  } catch(e) {
    var cleaned = '';
    var inStr = false, esc = false;
    for (var ci = 0; ci < raw.length; ci++) {
      var ch = raw[ci];
      if (esc)              { cleaned += ch; esc = false; continue; }
      if (ch === '\\' && inStr) { cleaned += ch; esc = true;  continue; }
      if (ch === '"')       { cleaned += ch; inStr = !inStr; continue; }
      if (inStr) {
        var code = ch.charCodeAt(0);
        if (ch === '\n') { cleaned += '\\n'; continue; }
        if (ch === '\r') { cleaned += '\\r'; continue; }
        if (ch === '\t') { cleaned += '\\t'; continue; }
        if (code < 0x20 || code === 0x7F) continue;
      }
      cleaned += ch;
    }
    parsed = JSON.parse(cleaned);
  }
  return parsed;
}

// ── Indexing job ──────────────────────────────────────────────────────────────

/**
 * RUN FROM EDITOR (or time trigger).
 * Scans CS_FOLDERS_, extracts text from each Google Slides file,
 * classifies as Case Study or Key Learning, tags + summarises with Gemini,
 * and appends to the hidden "CS & KL Index" sheet.
 * Safe to re-run — already-indexed URLs are skipped.
 */
function runCaseStudyIndexingJob() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('⏳ Another CS indexing run is still in progress. Skipping.');
    return;
  }

  var START_MS = new Date().getTime();
  var LIMIT_MS = 4 * 60 * 1000;

  var indexSheet  = getOrCreateCaseStudyIndexSheet_();
  var indexedUrls = {};
  var lastRow     = indexSheet.getLastRow();
  if (lastRow > 1) {
    indexSheet.getRange(2, 3, lastRow - 1, 1).getValues()
      .forEach(function(r) { if (r[0]) indexedUrls[r[0]] = true; });
  }

  // Build queue from all configured folders
  var queue = [];
  CS_FOLDERS_.forEach(function(cfg) {
    if (!cfg.folderId || cfg.folderId.indexOf('PASTE_') === 0) {
      Logger.log('⚠ Skipping year ' + cfg.year + ' — folder ID not configured.');
      return;
    }
    try {
      var folder = DriveApp.getFolderById(cfg.folderId);
      var files  = folder.getFilesByType(MimeType.GOOGLE_SLIDES);
      while (files.hasNext()) {
        var f   = files.next();
        var url = f.getUrl();
        if (!indexedUrls[url]) {
          queue.push({ year: cfg.year, name: f.getName(), url: url, fileId: f.getId() });
        }
      }
    } catch(e) {
      Logger.log('⚠ Could not access folder for year ' + cfg.year + ': ' + e.toString());
    }
  });

  if (queue.length === 0) {
    Logger.log('✅ All case studies / key learnings already indexed.');
    stopCaseStudyIndexingTrigger_();
    lock.releaseLock();
    return;
  }

  Logger.log('Queue: ' + queue.length + ' unindexed files. Processing up to ' + CS_BATCH_SIZE_ + ' this run.');

  var newRows   = [];
  var processed = 0;

  for (var i = 0; i < queue.length && processed < CS_BATCH_SIZE_; i++) {
    if (new Date().getTime() - START_MS > LIMIT_MS) {
      Logger.log('⏱ Time limit reached at ' + processed + ' files.');
      break;
    }

    var f = queue[i];
    Logger.log('[' + (processed + 1) + '/' + Math.min(queue.length, CS_BATCH_SIZE_) + '] ' + f.name);

    var text = extractTextFromUrl_(f.url, 15);
    if (!text) {
      Logger.log('  ⚠ Could not extract text — skipping.');
      // Still record it so we don't retry forever
      newRows.push([f.year, f.name, f.url, 'Unknown', '', '', '[]', 'Could not extract text.', new Date().toISOString()]);
      processed++;
      continue;
    }

    try {
      var result = classifyAndSummariseForCsIndex_(text, f.name);
      var type     = result.type     || 'Unknown';
      var brand    = result.brand    || '';
      var campaign = result.campaign || '';
      var tags     = Array.isArray(result.tags) ? result.tags : [];
      var summary  = result.summary  || '';
      Logger.log('  → ' + type + ' | ' + brand + ' | ' + tags.length + ' tags');
      newRows.push([f.year, f.name, f.url, type, brand, campaign, JSON.stringify(tags), summary, new Date().toISOString()]);
    } catch(e) {
      Logger.log('  ⚠ Gemini error: ' + e.toString());
      newRows.push([f.year, f.name, f.url, 'Unknown', '', '', '[]', '', new Date().toISOString()]);
    }

    processed++;
    Utilities.sleep(500); // gentle rate limiting
  }

  if (newRows.length > 0) {
    indexSheet.getRange(indexSheet.getLastRow() + 1, 1, newRows.length, 9).setValues(newRows);
    Logger.log('✅ Indexed ' + newRows.length + ' files this run.');
  }

  lock.releaseLock();
}

/** Run once from editor to start background indexing every minute. */
function setupCaseStudyIndexingTrigger() {
  stopCaseStudyIndexingTrigger_();
  ScriptApp.newTrigger('runCaseStudyIndexingJob').timeBased().everyMinutes(1).create();
  Logger.log('✅ CS indexing trigger created. Runs every minute until queue is empty.');
}

/** Run from editor to cancel. */
function stopCaseStudyIndexingTrigger() {
  stopCaseStudyIndexingTrigger_();
  Logger.log('CS indexing trigger removed.');
}

function stopCaseStudyIndexingTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runCaseStudyIndexingJob') ScriptApp.deleteTrigger(t);
  });
}

/** Run from editor to see progress. */
function getCaseStudyIndexStats() {
  var sheet = getOrCreateCaseStudyIndexSheet_();
  if (sheet.getLastRow() <= 1) { Logger.log('CS Index is empty. Run setupCaseStudyIndexingTrigger().'); return; }
  var count = sheet.getLastRow() - 1;
  var data  = sheet.getRange(2, 1, count, 4).getValues();
  var byType = {}, byYear = {};
  data.forEach(function(r) {
    var yr = r[0] || 'Unknown'; var tp = r[3] || 'Unknown';
    byType[tp] = (byType[tp] || 0) + 1;
    byYear[yr] = (byYear[yr] || 0) + 1;
  });
  Logger.log('📊 CS & KL Index: ' + count + ' total');
  for (var t in byType) Logger.log('   ' + t + ': ' + byType[t]);
  for (var y in byYear) Logger.log('   ' + y + ': ' + byYear[y]);
}

// ── Query function ────────────────────────────────────────────────────────────

/**
 * Called from the Strategist Dashboard when a strategist clicks "📚 Past Learnings".
 * Returns { casestudies: [...], keylearnings: [...] } — top 3 of each type
 * matched to the current brief by tags.
 */
function getRelevantCaseStudies(brand, campaign, industry, briefUrls) {
  try {
    if (typeof briefUrls === 'string') briefUrls = briefUrls ? [briefUrls] : [];
    briefUrls = (briefUrls || []).filter(Boolean);

    var cacheKey = 'cskl1_' + (brand + '|' + industry + '|' + briefUrls.join(','))
      .toLowerCase().replace(/[^a-z0-9|,]/g, '_').substring(0, 200);
    var cache  = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch(e) {}
    }

    var sheet   = getOrCreateCaseStudyIndexSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { casestudies: [], keylearnings: [], empty: true };

    // Tag the brief
    var briefText = extractBriefText_(briefUrls, 10, 1500);
    var tagInput  =
      (briefText ? briefText + '\n\n' : '') +
      'Brand: ' + brand + '\nCampaign: ' + campaign +
      '\nIndustry: ' + (industry || '');
    var briefTags = tagWithGemini_(tagInput, brand, campaign, industry, '');
    if (!briefTags.length) return { casestudies: [], keylearnings: [], empty: true };

    // Score every row
    var rows   = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var csArr  = [];
    var klArr  = [];

    rows.forEach(function(row) {
      var year     = String(row[0] || '');
      var name     = String(row[1] || '');
      var url      = String(row[2] || '');
      var type     = String(row[3] || '');
      var brand_   = String(row[4] || '');
      var campaign_= String(row[5] || '');
      var summary  = String(row[7] || '');
      var itemTags = [];
      try { itemTags = JSON.parse(String(row[6] || '[]')); } catch(e) {}

      if (!url || !itemTags.length || type === 'Unknown') return;

      var score     = computeTagScore_(briefTags, itemTags);
      if (score === 0) return;
      var matchTags = getMatchingTagNames_(briefTags, itemTags);

      var entry = { year: year, name: name, url: url, brand: brand_,
                    campaign: campaign_, summary: summary,
                    score: score, matchTags: matchTags };

      if (type === 'Case Study')   csArr.push(entry);
      else if (type === 'Key Learning') klArr.push(entry);
    });

    csArr.sort(function(a, b) { return b.score - a.score; });
    klArr.sort(function(a, b) { return b.score - a.score; });

    var result = {
      casestudies:  csArr.slice(0, 3),
      keylearnings: klArr.slice(0, 3),
      empty: csArr.length === 0 && klArr.length === 0
    };

    putCacheWithRegistry_(cache, cacheKey, JSON.stringify(result), 21600);
    return result;

  } catch(e) {
    Logger.log('getRelevantCaseStudies error: ' + e.toString());
    return { casestudies: [], keylearnings: [], empty: true, error: e.message };
  }
}

/**
 * Called after renderLearnings — re-reads top 3 CS + top 3 KL files with brief context
 * and generates brief-aware insights to replace the generic indexed summaries.
 *
 * @param {Object[]} csResults   Top CS entries from getRelevantCaseStudies
 * @param {Object[]} klResults   Top KL entries from getRelevantCaseStudies
 * @param {string}   brand
 * @param {string}   campaign
 * @param {string}   industry
 * @param {string[]} briefUrls
 * @returns {{ casestudies: [{index,insight}], keylearnings: [{index,insight}] }}
 */
function getCaseStudyInsights(csResults, klResults, brand, campaign, industry, briefUrls) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) throw new Error('GEMINI_API_KEY not set');

    if (typeof briefUrls === 'string') briefUrls = briefUrls ? [briefUrls] : [];
    briefUrls = (briefUrls || []).filter(Boolean);

    // Cache key based on top URLs + brief
    var topCs = (csResults || []).slice(0, 3);
    var topKl = (klResults || []).slice(0, 3);
    var cacheKey = 'csins1_' + (
      topCs.map(function(r) { return r.url; }).join('|') + '||' +
      topKl.map(function(r) { return r.url; }).join('|') + '||' +
      brand + '|' + industry
    ).replace(/[^a-z0-9|]/gi, '_').substring(0, 200);

    var cache  = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch(e) {}
    }

    // Extract brief context
    var briefText = extractBriefText_(briefUrls, 10, 1500);
    var briefSection = briefText
      ? 'CURRENT BRIEF (from brief files):\n' + briefText.substring(0, 2000)
      : 'CURRENT BRIEF:\nBrand: ' + brand + '\nCampaign: ' + campaign + '\nIndustry: ' + (industry || 'N/A');

    // Extract text from top CS files
    var csBlocks = topCs.map(function(r, i) {
      var text = extractTextFromUrl_(r.url, 15);
      return '--- CASE STUDY ' + (i + 1) + ': ' + (r.brand || '') + ' / ' + (r.campaign || r.name || '') + ' [' + (r.year || '') + '] ---\n' +
        (text ? text.substring(0, 1500) : '[could not extract text]');
    }).join('\n\n');

    // Extract text from top KL files
    var klBlocks = topKl.map(function(r, i) {
      var text = extractTextFromUrl_(r.url, 15);
      return '--- KEY LEARNING ' + (i + 1) + ': ' + (r.brand || '') + ' / ' + (r.campaign || r.name || '') + ' [' + (r.year || '') + '] ---\n' +
        (text ? text.substring(0, 1500) : '[could not extract text]');
    }).join('\n\n');

    var prompt =
      'You are reviewing past campaign documents for a strategist who is currently working on a new brief. ' +
      'Your job is to write brief-aware insights — not generic summaries, but specific connections between each document and the current brief.\n\n' +
      briefSection + '\n\n' +
      '════ CASE STUDIES ════\n' + (csBlocks || '[none]') + '\n\n' +
      '════ KEY LEARNINGS ════\n' + (klBlocks || '[none]') + '\n\n' +
      'For each document, write exactly 2 sentences grounded in both the document AND the current brief:\n\n' +
      'FOR EACH CASE STUDY:\n' +
      '  Persona: Senior Campaign Strategist who reads past work to find what is replicable RIGHT NOW.\n' +
      '  Sentence 1 — WHAT WORKED: The single most specific creative idea or mechanic from this case study. ' +
      'Name it precisely — no generic phrases like "engaged audiences" or "leveraged social media".\n' +
      '  Sentence 2 — STEAL THIS: The one element from this case study that directly applies to the current brief. ' +
      'Be explicit about the connection — "Given [brief context], adapt [specific element] by [how]." ' +
      'If nothing maps cleanly, say so honestly.\n\n' +
      'FOR EACH KEY LEARNING:\n' +
      '  Persona: Tough Campaign Director reviewing a post-mortem. No softening, no corporate language.\n' +
      '  Sentence 1 — WHAT FAILED: The specific platform, format, or assumption that underperformed, with a metric if available.\n' +
      '  Sentence 2 — WATCH OUT: Given the current brief, name the exact risk the strategist should pre-empt. ' +
      'Start with a verb: "Avoid...", "Test...", "Do not assume...", "Cap spend on..." etc.\n\n' +
      'Rules:\n' +
      '- Maximum 60 words per document (both sentences combined) — keep it tight but never cut mid-sentence\n' +
      '- Every sentence must reference real content from both the document and the brief\n' +
      '- Zero filler: no "compelling", "impactful", "innovative", "strong results"\n' +
      '- If a document could not be extracted, write: "Could not read this document — open it directly."\n\n' +
      'Return ONLY a JSON object, no markdown. You MUST include exactly one entry for EVERY document listed above — ' +
      'do not skip any, even if content was sparse. The arrays must have exactly as many items as documents provided:\n' +
      '{\n' +
      '  "casestudies": [{"index":0,"insight":"..."},{"index":1,"insight":"..."},{"index":2,"insight":"..."}],\n' +
      '  "keylearnings": [{"index":0,"insight":"..."},{"index":1,"insight":"..."},{"index":2,"insight":"..."}]\n' +
      '}';

    var resp = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.7, maxOutputTokens: 4096 }
        }),
        muteHttpExceptions: true
      }
    );

    var json = JSON.parse(resp.getContentText());
    if (json.error) throw new Error('Gemini: ' + json.error.message);

    var allParts = (json.candidates[0].content.parts || []);
    console.log('🔍 getCaseStudyInsights: parts count=' + allParts.length +
      ' finishReason=' + (json.candidates[0].finishReason || 'none'));

    var raw = allParts
      .filter(function(p) { return p.text && !p.thought; })
      .map(function(p) { return p.text; }).join('').trim()
      .replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/i, '').trim();

    console.log('🔍 getCaseStudyInsights: raw length=' + raw.length + ' preview=' + raw.substring(0, 100));

    // Robust JSON parse (char-walker)
    var parsed;
    try {
      parsed = JSON.parse(raw);
    } catch(e) {
      var cleaned = ''; var inStr = false; var esc = false;
      for (var ci = 0; ci < raw.length; ci++) {
        var ch = raw[ci];
        if (esc)              { cleaned += ch; esc = false; continue; }
        if (ch === '\\' && inStr) { cleaned += ch; esc = true; continue; }
        if (ch === '"')       { cleaned += ch; inStr = !inStr; continue; }
        if (inStr) {
          var code = ch.charCodeAt(0);
          if (ch === '\n') { cleaned += '\\n'; continue; }
          if (ch === '\r') { cleaned += '\\r'; continue; }
          if (ch === '\t') { cleaned += '\\t'; continue; }
          if (code < 0x20 || code === 0x7F) continue;
        }
        cleaned += ch;
      }
      parsed = JSON.parse(cleaned);
    }

    putCacheWithRegistry_(cache, cacheKey, JSON.stringify(parsed), 21600);
    return parsed;

  } catch(e) {
    Logger.log('getCaseStudyInsights error: ' + e.toString());
    return { casestudies: [], keylearnings: [] };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// CAMPAIGN INTEL — Brand Performance Explorer
// Architecture: Option C (gviz query) per CampaignIntel_Architecture_Decision.md
//
// SETUP STEPS:
//   1. Paste the Campaign Encyclopedia Google Spreadsheet ID into CAMPAIGN_ENC_SS_ID_
//   2. Confirm the sheet tab names in CAMPAIGN_ENC_SHEETS_ match your file
//   3. Run runCampaignIndexJob() once manually to build the slim index
//   4. Set up a daily time-trigger pointing to runCampaignIndexJob()
//   5. Navigate to ?page=brandintel to open Brand Intel
// ═══════════════════════════════════════════════════════════════════════════════

var CAMPAIGN_ENC_SS_ID_    = '163wGwCJyCRO_CIT5kXuazmXMvYQUWKqtEgIX79JUa8A';
var CAMPAIGN_INDEX_SHEET_  = 'Campaign Data Index';

// Each entry can include an optional `colOverride` map to handle sheets that
// used a different column layout before the 2025/2026 template was standardised.
// Any key in colOverride replaces the matching key in CE_COLS_ for that sheet only.
// Set a key to null to mark it as unpopulated (will write empty string for that col).
var CAMPAIGN_ENC_SHEETS_   = [
  {
    year: '2023',
    sheetName: 'Campaign Encyclopedia 2023',
    // 2023 pre-standardisation layout:
    //   AS (44) = Reach (article page views)   ← swapped vs 2025/2026
    //   AQ (42) = Delivery% ratio              ← swapped vs 2025/2026
    //   AO (40) = not used / empty             ← no ER% for article rows
    colOverride: { reach: 44, deliveryPct: 42, erPct: null }
  },
  {
    year: '2024',
    sheetName: 'Campaign Encyclopedia 2024'
    // Assumed to follow the standardised template (same as 2025/2026).
    // If numbers look wrong after indexing, add a colOverride here.
  },
  {
    year: '2025',
    sheetName: 'Campaign Encyclopedia 2025/2026'
    // Standardised template confirmed via row-3 headers. Year per row derived from Start Date.
  },
];

// Column indices in raw Campaign Encyclopedia rows (0-based, data starts at row 4)
// Confirmed against row 3 headers of the 2025/2026 sheet:
//   AO = "Delivery %"  (Story Stats group)
//   AQ = FB Reach / IG Reach / TT Video Views / YT Impressions / PCTO Est. PVs
//   AR = FB/IG/TT Engagements / PCTO Clicks
//   AS = FB ER (%) / IG ER (%) / TT ER (%) / Banner CTR (%) / PCTO CTR (%)
// NOTE: 2023 rows were entered before this template was standardised — their
// AQ/AS values may be swapped vs 2025/2026. Filter by Year to get clean data.
var CE_COLS_ = {
  campaignId:  0,   // A — Campaign ID
  campaignName: 2,  // C — Campaign Name
  status:      3,   // D — Status
  startDate:   13,  // N — Start Date
  brand:       17,  // R — Brand
  category:    20,  // U — Category
  platform:    21,  // V — Platform (REV Media IP / page)
  type:        22,  // W — Type (content type + channel)
  landingPage: 26,  // AA — Landing Page / "View Content" link
  deliveryPct: 40,  // AO — Delivery % ratio (1.074 = 107.4% delivered)
  reach:       42,  // AQ — Reach (FB/IG/TT/YT/article page views)
  engagements: 43,  // AR — Engagements (likes + comments + shares)
  erPct:       44,  // AS — ER % (stored as percentage: 1.68 = 1.68%)
};
var CE_READ_WIDTH_ = 45;  // Read cols A–AS (1–45)

// Content types that are internal cost entries / non-organic — exclude from index
var CE_EXCLUDE_TYPES_ = [
  'REV Sponsored Production Cost Savings',
  'Video Ads'
];

/**
 * Full-rebuild sync job: reads every year tab of the Campaign Encyclopedia,
 * filters and maps to a 14-column slim schema, and writes everything to the
 * hidden "Campaign Data Index" tab in this spreadsheet.
 *
 * Run ONCE manually to seed, then set a daily time-trigger on this function.
 * Full rebuild is intentional — many rows get their metrics updated post-delivery
 * and append-only would serve stale numbers.
 */
function runCampaignIndexJob() {
  try {
    var encSs  = SpreadsheetApp.openById(CAMPAIGN_ENC_SS_ID_);
    var mainSs = SpreadsheetApp.getActiveSpreadsheet();

    // Get or create the slim index sheet (hidden)
    var indexSheet = mainSs.getSheetByName(CAMPAIGN_INDEX_SHEET_);
    if (!indexSheet) {
      indexSheet = mainSs.insertSheet(CAMPAIGN_INDEX_SHEET_);
      indexSheet.hideSheet();
    } else {
      indexSheet.clearContents();
    }

    // Slim schema headers (A–N = 14 cols)
    var headers = [
      'Campaign ID',  // A
      'Campaign Name',// B
      'Brand',        // C  ← primary filter
      'Category',     // D  ← filter
      'Platform',     // E  ← filter (REV Media IP)
      'Type',         // F  ← filter (content type + channel)
      'Year',         // G  ← filter
      'Start Date',   // H
      'Status',       // I
      'ER%',          // J  ← primary sort/metric
      'Reach',        // K
      'Engagements',  // L
      'Delivery%',    // M  (ratio: 1.0 = 100%, >1 over-delivered, <1 under)
      'Landing Page'  // N  (view content link, ~65% populated)
    ];
    indexSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    indexSheet.setFrozenRows(1);

    // Build the exclusion set for fast lookup
    var excludeSet = {};
    CE_EXCLUDE_TYPES_.forEach(function(t) { excludeSet[t.toLowerCase()] = true; });

    var allRows = [];

    CAMPAIGN_ENC_SHEETS_.forEach(function(cfg) {
      try {
        var srcSheet = encSs.getSheetByName(cfg.sheetName);
        if (!srcSheet) {
          Logger.log('⚠ Sheet not found: "' + cfg.sheetName + '" — check CAMPAIGN_ENC_SHEETS_ tab names');
          return;
        }

        var lastRow = srcSheet.getLastRow();
        // Rows 1–3 are title / group-header / column-header; data starts at row 4
        if (lastRow < 4) {
          Logger.log('⚠ Sheet "' + cfg.sheetName + '" has no data rows.');
          return;
        }

        // Merge per-sheet column overrides into the base CE_COLS_ map.
        // colOverride keys replace the base value; null means "column not present".
        var cols = {};
        for (var k in CE_COLS_) { cols[k] = CE_COLS_[k]; }
        if (cfg.colOverride) {
          for (var ok in cfg.colOverride) { cols[ok] = cfg.colOverride[ok]; }
        }

        var numDataRows = lastRow - 3;
        var data = srcSheet.getRange(4, 1, numDataRows, CE_READ_WIDTH_).getValues();
        var rowsBefore = allRows.length;

        data.forEach(function(row) {
          var campaignId = String(row[cols.campaignId] || '').trim();
          if (!campaignId) return;   // skip completely empty rows

          var type = String(row[cols.type] || '').trim();
          if (excludeSet[type.toLowerCase()]) return;  // skip excluded types

          // Derive Year from Start Date; fallback to Campaign ID prefix (e.g. JUN26 → 2026)
          var startDate = row[cols.startDate];
          var year = '';
          var startDateStr = '';
          if (startDate instanceof Date && !isNaN(startDate.getTime())) {
            year = String(startDate.getFullYear());
            startDateStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');
          } else {
            // Fallback: "JUN26-001A" → first 3 letters + 2 digits → year = 20YY
            var idMatch = campaignId.match(/[A-Z]{3}(\d{2})/i);
            if (idMatch) {
              var yy = parseInt(idMatch[1]);
              year = String(yy < 50 ? 2000 + yy : 1900 + yy);
            }
          }

          function colVal(key) {
            var idx = cols[key];
            if (idx === null || idx === undefined) return '';
            var v = row[idx];
            return (v !== '' && v !== null && v !== undefined) ? v : '';
          }

          allRows.push([
            campaignId,
            String(row[cols.campaignName] || '').trim(),
            String(row[cols.brand]        || '').trim(),
            String(row[cols.category]     || '').trim(),
            String(row[cols.platform]     || '').trim(),
            type,
            year,
            startDateStr,
            String(row[cols.status]       || '').trim(),
            colVal('erPct'),
            colVal('reach'),
            colVal('engagements'),
            colVal('deliveryPct'),
            String(row[cols.landingPage]  || '').trim(),
          ]);
        });

        Logger.log('✓ "' + cfg.sheetName + '": read ' + data.length + ' rows → ' +
          (allRows.length - rowsBefore) + ' kept' +
          (cfg.colOverride ? ' [col override applied]' : ''));

      } catch(sheetErr) {
        Logger.log('⚠ Error reading "' + cfg.sheetName + '": ' + sheetErr.toString());
      }
    });

    // Single bulk write — much faster than per-row appends
    if (allRows.length > 0) {
      indexSheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
    }

    Logger.log('✅ runCampaignIndexJob complete — ' + allRows.length +
      ' rows written to "' + CAMPAIGN_INDEX_SHEET_ + '"');

  } catch(e) {
    Logger.log('❌ runCampaignIndexJob error: ' + e.toString());
    throw e;
  }
}

/**
 * Returns meta about the Campaign Data Index sheet so the frontend can
 * build a gviz query URL without any backend round-trip per query.
 */
function getCampaignIndexInfo() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CAMPAIGN_INDEX_SHEET_);
  var rowCount = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
  return {
    spreadsheetId: ss.getId(),
    sheetName:     CAMPAIGN_INDEX_SHEET_,
    exists:        !!sheet,
    rowCount:      rowCount
  };
}

/**
 * Server-side fallback for queryCampaignData — used only when the client-side
 * gviz fetch fails (CORS / auth). Filters the slim index in Apps Script and
 * returns matching rows.
 *
 * Filters object keys: brand, category, platform, type, year, metrics
 * ("yes" in metrics = only rows where ER% is populated)
 */
function queryCampaignData(filters) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CAMPAIGN_INDEX_SHEET_);
    if (!sheet || sheet.getLastRow() < 2) return { rows: [] };

    filters = filters || {};
    var brand    = String(filters.brand    || '').trim().toLowerCase();
    var category = String(filters.category || '').trim().toLowerCase();
    var platform = String(filters.platform || '').trim().toLowerCase();
    var type     = String(filters.type     || '').trim().toLowerCase();
    var year     = String(filters.year     || '').trim();
    var metrics  = String(filters.metrics  || '');

    var numRows = sheet.getLastRow() - 1;
    // Read all 14 cols
    var data = sheet.getRange(2, 1, numRows, 14).getValues();

    var results = [];
    data.forEach(function(row) {
      if (brand    && String(row[2]  || '').toLowerCase() !== brand)    return;
      if (category && String(row[3]  || '').toLowerCase() !== category) return;
      if (platform && String(row[4]  || '').toLowerCase() !== platform) return;
      if (type     && String(row[5]  || '').toLowerCase() !== type)     return;
      if (year     && String(row[6]  || '') !== year)                   return;
      if (metrics === 'yes' && (row[9] === '' || row[9] === null))      return;

      results.push({
        id:          String(row[0]  || ''),
        name:        String(row[1]  || ''),
        brand:       String(row[2]  || ''),
        category:    String(row[3]  || ''),
        platform:    String(row[4]  || ''),
        type:        String(row[5]  || ''),
        year:        String(row[6]  || ''),
        date:        String(row[7]  || ''),
        status:      String(row[8]  || ''),
        er:          row[9]  !== '' && row[9]  !== null ? row[9]  : '',
        reach:       row[10] !== '' && row[10] !== null ? row[10] : '',
        engagements: row[11] !== '' && row[11] !== null ? row[11] : '',
        dlv:         row[12] !== '' && row[12] !== null ? row[12] : '',
        link:        String(row[13] || '')
      });
    });

    // Sort by ER% descending (backend default; frontend will re-sort)
    results.sort(function(a, b) {
      return (parseFloat(b.er) || -999) - (parseFloat(a.er) || -999);
    });

    Logger.log('queryCampaignData: ' + results.length + ' rows matched');
    return { rows: results };

  } catch(e) {
    Logger.log('queryCampaignData error: ' + e.toString());
    return { rows: [], error: e.message };
  }
}

/**
 * Reads distinct values from the slim index for each filter dimension.
 * Called once on Brand Intel page load to populate the filter dropdowns.
 * Only reads 5 columns (C–G), so it's fast even at 46k rows.
 */
function getCampaignFilterOptions() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CAMPAIGN_INDEX_SHEET_);
  if (!sheet || sheet.getLastRow() < 2) {
    return { brands: [], categories: [], platforms: [], types: [], years: [],
             rowCount: 0 };
  }

  var numRows = sheet.getLastRow() - 1;
  // Cols C=Brand(3), D=Category(4), E=Platform(5), F=Type(6), G=Year(7) → read 5 cols starting at col 3
  var data = sheet.getRange(2, 3, numRows, 5).getValues();

  var brands = {}, categories = {}, platforms = {}, types = {}, years = {};
  data.forEach(function(row) {
    var b = String(row[0] || '').trim();
    var c = String(row[1] || '').trim();
    var p = String(row[2] || '').trim();
    var t = String(row[3] || '').trim();
    var y = String(row[4] || '').trim();
    if (b && b !== '-') brands[b]      = true;
    if (c && c !== '-') categories[c]  = true;
    if (p && p !== '-') platforms[p]   = true;
    if (t && t !== '-') types[t]       = true;
    if (y)              years[y]       = true;
  });

  function sorted(obj) { return Object.keys(obj).sort(); }
  return {
    brands:     sorted(brands),
    categories: sorted(categories),
    platforms:  sorted(platforms),
    types:      sorted(types),
    years:      sorted(years).reverse(),   // newest first
    rowCount:   numRows
  };
}

/**
 * REV Intel — Format Performance Benchmarks
 * Feeds the brief context to Gemini (REV Intel persona), extracts which content
 * formats are implied, maps them to the live campaign taxonomy, and returns
 * query groups the client can fire against the Campaign Data Index via gviz.
 *
 * @param {string} brand
 * @param {string} campaign
 * @param {string} industry
 * @param {string} briefSummary  (first ~300 chars is enough)
 * @param {string} deckType      ("Standard" | "KLR Deck" | etc.)
 * @return {{ success, queries, ssId, sheetName } | { success, error }}
 */
function getBriefFormatIntel(brand, campaign, industry, briefSummary, deckType) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CAMPAIGN_INDEX_SHEET_);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, error: 'Campaign index not built yet. Run runCampaignIndexJob() first.' };
    }

    // ── Cache (v3 key — busts old cache to include median3) ───────────────
    var cacheKey = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      'fintel3|' + (brand || '') + '|' + (campaign || '') + '|' + (industry || '') + '|' +
      (briefSummary || '').substring(0, 100)
    ).reduce(function(acc, b) {
      return acc + ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
    }, '');
    cacheKey = 'fintel3_' + cacheKey;

    var cache  = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) {
      Logger.log('⚡ getBriefFormatIntel: cache hit');
      return JSON.parse(cached);
    }

    // ── Get live taxonomy from the index ─────────────────────────────────
    var opts = getCampaignFilterOptions();
    if (!opts.types.length) {
      return { success: false, error: 'Campaign index is empty. Run runCampaignIndexJob() first.' };
    }

    // ── Build Gemini prompt ───────────────────────────────────────────────
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) throw new Error('GEMINI_API_KEY not set in Script Properties');

    var prompt =
      'You are REV Intel, the campaign intelligence engine for REV Media Group\'s Strategy Department in Malaysia. ' +
      'You have planned hundreds of digital content campaigns across REV\'s owned platforms — SAYS, OHBULAN!, Mashable SEA, Vocket, SirapLimau, ViralCham — and their social channels. ' +
      'You speak the language of Malaysian digital media strategy fluently.\n\n' +

      'A strategist has received a new client brief. Read it and identify which content formats are either explicitly mentioned or strongly implied by the objectives, audience strategy, or deliverables described. ' +
      'Then map each format to the EXACT Type values used in REV\'s campaign tracker. Only use values from the Valid Types list below — do not invent new ones.\n\n' +

      'FORMAT MAPPING RULES (these are the most common — apply judgement for anything not listed):\n' +
      '• "Short-form video / reels / viral content / social-first" → look for: IG Reel, TT Post, FB Reel\n' +
      '• "Social media posts / social content / social amplification" → look for: IG Post, FB Post, TT Post\n' +
      '• "Editorial / branded content / storytelling / native content / article" → look for: Spotlight Story\n' +
      '• "YouTube / long-form video / documentary" → look for: YT Upload, YT Video\n' +
      '• "Gallery / carousel / swipe" → look for: Gallery Post, FB Carousel\n' +
      '• "Banner / programmatic / display / digital ads" → look for: Banner\n' +
      '• If the brief is vague, infer from industry + objective (e.g. F&B + engagement = social posts)\n\n' +

      'BRIEF CONTEXT:\n' +
      'Brand: '         + (brand        || 'Unknown') + '\n' +
      'Campaign: '      + (campaign     || 'Unknown') + '\n' +
      'Industry: '      + (industry     || 'Unknown') + '\n' +
      'Deck Type: '     + (deckType     || 'Standard') + '\n' +
      'Brief Summary: ' + (briefSummary || '(no summary provided)') + '\n\n' +

      'VALID TYPE VALUES (use EXACT spelling — case-sensitive):\n' +
      opts.types.join(', ') + '\n\n' +

      'VALID CATEGORY VALUES (use EXACT spelling):\n' +
      opts.categories.join(', ') + '\n\n' +

      'Return up to 4 query groups, ranked by how likely this format appears in the final proposal. ' +
      'Group related formats together (e.g. IG Reel + TT Post = one "Short-form Video" group).\n\n' +

      'Return ONLY valid JSON — no markdown, no code fences, no extra text:\n' +
      '{"queries":[{"label":"Short-form Social Video","types":["IG Reel","TT Post"],"category":"F&B",' +
      '"rationale":"Brief mentions a social-first approach targeting Gen Z — maps to IG Reel and TT Post"}]}';

    var resp = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
      {
        method:          'post',
        contentType:     'application/json',
        payload:         JSON.stringify({
          contents:         [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.3, maxOutputTokens: 8192 }
        }),
        muteHttpExceptions: true
      }
    );

    var rJson = JSON.parse(resp.getContentText());
    if (rJson.error) throw new Error('Gemini API: ' + rJson.error.message);

    var parts = (((rJson.candidates || [])[0] || {}).content || {}).parts || [];
    var text  = parts
      .filter(function(p) { return p.text && !p.thought; })
      .map(function(p) { return p.text; })
      .join('').trim();

    var jsonMatch = text.match(/\{[\s\S]*\}/);
    var parsed;
    try { parsed = JSON.parse(jsonMatch ? jsonMatch[0] : text); }
    catch(e) {
      Logger.log('❌ getBriefFormatIntel JSON parse failed: ' + text.substring(0, 300));
      return { success: false, error: 'Could not parse Gemini response. Try again.' };
    }

    // ── Validate types + category against live taxonomy ───────────────────
    var typeMap = {};
    opts.types.forEach(function(t) { typeMap[t.toLowerCase()] = t; });
    var catMap = {};
    opts.categories.forEach(function(c) { catMap[c.toLowerCase()] = c; });

    var validated = (parsed.queries || []).map(function(q) {
      var cleanTypes = (q.types || [])
        .map(function(t) { return typeMap[String(t).toLowerCase()] || null; })
        .filter(Boolean);
      if (!cleanTypes.length) return null;
      return {
        label:     String(q.label     || '').trim(),
        types:     cleanTypes,
        category:  catMap[String(q.category || '').toLowerCase()] || '',
        rationale: String(q.rationale || '').trim()
      };
    }).filter(Boolean).slice(0, 4);

    Logger.log('getBriefFormatIntel: ' + validated.length + ' valid query group(s)');
    if (!validated.length) {
      return { success: true, groups: [] };
    }

    // ── Query the index server-side (one read, filter N times) ───────────
    var numRows = sheet.getLastRow() - 1;
    var data    = sheet.getRange(2, 1, numRows, 14).getValues();
    // Index cols (0-based): B=1 Name, C=2 Brand, D=3 Category, F=5 Type, J=9 ER%, N=13 Link

    var groups = validated.map(function(q) {
      var typeSet  = {};
      q.types.forEach(function(t) { typeSet[t] = true; });
      var catLower = q.category ? q.category.toLowerCase() : '';

      function filterRows(useCat) {
        var out = [];
        data.forEach(function(row) {
          var type = String(row[5] || '').trim();
          var cat  = String(row[3] || '').trim();
          if (!typeSet[type]) return;
          if (useCat && catLower && cat.toLowerCase() !== catLower) return;
          out.push({
            name:  String(row[1] || ''),
            brand: String(row[2] || ''),
            type:  type,
            er:    (row[9] !== '' && row[9] !== null && row[9] !== undefined) ? row[9] : null,
            link:  String(row[13] || '')
          });
        });
        return out;
      }

      var rows = filterRows(true);
      if (!rows.length && catLower) rows = filterRows(false); // fallback: drop category

      // Sort by ER% desc; compute avg from rows that have ER%
      var withEr = rows
        .filter(function(r) { return r.er !== null && !isNaN(parseFloat(r.er)) && parseFloat(r.er) > 0; })
        .sort(function(a, b) { return parseFloat(b.er) - parseFloat(a.er); });

      var avgEr = withEr.length
        ? Math.round((withEr.reduce(function(s, r) { return s + parseFloat(r.er); }, 0) / withEr.length) * 100) / 100
        : null;

      // Pick performers — one entry per brand (best ER% per brand)
      var srcRows = withEr.length ? withEr : rows;
      var seenBrands = {};
      var uniqueBrandsList = [];
      srcRows.forEach(function(r) {
        var b = (r.brand || r.name || '').toLowerCase();
        if (b && seenBrands[b]) return;
        if (b) seenBrands[b] = true;
        uniqueBrandsList.push({ brand: r.brand || r.name, type: r.type, er: r.er, link: r.link });
      });
      
      var top3 = uniqueBrandsList.slice(0, 3);
      
      // Calculate median3 based on the unique brands list
      var median3 = [];
      if (uniqueBrandsList.length > 6) {
        var midIdx = Math.floor(uniqueBrandsList.length / 2);
        var startIdx = Math.max(3, midIdx - 1);
        median3 = uniqueBrandsList.slice(startIdx, startIdx + 3);
      } else if (uniqueBrandsList.length > 3) {
        median3 = uniqueBrandsList.slice(3, 6);
      }

      Logger.log('  [' + q.label + '] total=' + rows.length + ' withEr=' + withEr.length + ' avgEr=' + avgEr);

      return {
        label:     q.label,
        rationale: q.rationale,
        types:     q.types,
        count:     rows.length,
        erCount:   withEr.length,
        avgEr:     avgEr,
        top3:      top3,
        median3:   median3
      };
    });

    var result = { success: true, groups: groups };
    putCacheWithRegistry_(cache, cacheKey, JSON.stringify(result), 21600); // 6-hr cache
    return result;

  } catch(e) {
    Logger.log('❌ getBriefFormatIntel error: ' + e.toString());
    return { success: false, error: e.message };
  }
}
