/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * README: HOW TO ENABLE GOOGLE SHEETS INTEGRATION
 *
 * This function sends data to a Google Sheet using a Google Apps Script Web App.
 * This is a secure method that avoids exposing any API keys or credentials on the frontend.
 *
 * --- SETUP INSTRUCTIONS (5 minutes) ---
 *
 * 1.  **GO TO GOOGLE APPS SCRIPT:**
 *     - Visit script.google.com and create a new project.
 *
 * 2.  **CREATE A GOOGLE APPS SCRIPT:**
 *     - A new script editor tab will open. Delete the default `function myFunction() {}` code.
 *     - **Copy and paste the entire script code provided below** into the editor.
 *     - The script is now set up to write to the specific Sheet ID you provided: '1NxJ_ZyV1KRIsQXFF7IcNWRg1VIBWRS0jd9ilXu0H2cs'
 *
 * 3.  **DEPLOY THE SCRIPT:**
 *     - At the top right of the script editor, click the blue "Deploy" button.
 *     - Select "New deployment".
 *     - For "Select type", click the gear icon and choose "Web app".
 *     - In the "Description" field, you can type "Valuation Reports API".
 *     - For "Execute as", select "Me (your.email@gmail.com)".
 *     - For "Who has access", select "Anyone". This is required for the web app to be callable.
 *     - Click "Deploy".
 *
 * 4.  **AUTHORIZE THE SCRIPT:**
 *     - Google will ask for permission to access your Google Sheets.
 *     - Click "Authorize access". Choose your Google account.
 *     - You might see a "Google hasn't verified this app" warning. This is normal. Click "Advanced", then "Go to [Your Script Name] (unsafe)".
 *     - Click "Allow" to grant the permissions.
 *
 * 5.  **GET THE WEB APP URL:**
 *     - After deploying, you will get a "Web app URL". Copy this URL.
 *
 * 6.  **UPDATE THE CODE:**
 *     - Paste the copied Web app URL into the `APPS_SCRIPT_URL` constant below, replacing the placeholder.
 *
 * --- APPS SCRIPT CODE TO PASTE (Step 2) ---
 *
 * // The ID of the specific Google Sheet you want to write to.
 * const TARGET_SHEET_ID = '1NxJ_ZyV1KRIsQXFF7IcNWRg1VIBWRS0jd9ilXu0H2cs';
 *
 * function doPost(e) {
 *   try {
 *     // Open the specific spreadsheet by its ID
 *     const spreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID);
 *     // Get the first sheet in the spreadsheet. You can also get it by name e.g. spreadsheet.getSheetByName("تقارير");
 *     const sheet = spreadsheet.getSheets()[0];
 *
 *     const data = JSON.parse(e.postData.contents);
 *     const headersMap = data.headersMap;
 *     const reports = data.reports;
 *
 *     const orderedKeys = Object.keys(headersMap);
 *     const orderedHeaderNames = orderedKeys.map(key => headersMap[key]);
 *
 *     // If sheet is empty, add headers
 *     if (sheet.getLastRow() === 0) {
 *       sheet.appendRow(orderedHeaderNames);
 *     }
 *
 *     // Add report data
 *     reports.forEach(report => {
 *       const row = orderedKeys.map(key => report[key] || 'غير موجود');
 *       sheet.appendRow(row);
 *     });
 *
 *     return ContentService.createTextOutput(JSON.stringify({
 *       status: 'success',
 *       message: `Successfully added ${reports.length} rows.`
 *     })).setMimeType(ContentService.MimeType.JSON);
 *
 *   } catch (error) {
 *     return ContentService.createTextOutput(JSON.stringify({
 *       status: 'error',
 *       message: 'Failed to write to sheet. Error: ' + error.toString() + ' | Ensure TARGET_SHEET_ID is correct and you have edit access.'
 *     })).setMimeType(ContentService.MimeType.JSON);
 *   }
 * }
 */

// STEP 6: PASTE YOUR WEB APP URL HERE
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz-n7C7Grw8twVpOY4EWIGpFW3IAQInCBfzNSSJaeUQJsq2GXQaD2BO_FfFZgCtqZfK7g/exec'; // هذا هو الرابط الصحيح

/**
 * Sends the extracted report data to a Google Apps Script endpoint which then writes to a Google Sheet.
 * @param reports An array of property report objects. For sheet upload, image fields are simplified to 'موجود'/'غير موجود'.
 * @param headersMap An object mapping property keys to translated header names.
 * @returns A promise that resolves with a status message for the UI.
 */
export async function uploadDataToGoogleSheet(reports, headersMap) {
    // قم بإزالة الجزء الزائد من الشرط هنا
    if (!APPS_SCRIPT_URL) { // فقط تحقق مما إذا كان الرابط فارغًا
        const message = "لم يتم تكوين Google Sheets. يرجى اتباع التعليمات في ملف google-api-handler.ts.";
        console.warn(message);
        return message;
    }

    const payload = {
        headersMap: headersMap,
        reports: reports,
    };

    try {
        const response = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            mode: 'no-cors', // Required for simple POST requests to Apps Script
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(payload),
            redirect: 'follow'
        });

        // Due to no-cors, we can't inspect the response. We will assume success if the request doesn't throw an error.
        const message = `تم إرسال ${reports.length} تقارير بنجاح إلى Google Sheets.`;
        console.log("Data sent to Google Sheets backend.");
        return message;

    } catch (error) {
        console.error("Failed to send data to Google Sheets:", error);
        return `فشل الاتصال بالخادم لإرسال البيانات إلى Google Sheets. الخطأ: ${error instanceof Error ? error.message : String(error)}`;
    }
}