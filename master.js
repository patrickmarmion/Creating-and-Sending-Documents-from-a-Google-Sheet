const scriptProperties = PropertiesService.getScriptProperties();
const APIKey = scriptProperties.getProperty('apiKey') || null;
const sandboxAPIKey = scriptProperties.getProperty('sandboxApiKey') || null;
const templateId = scriptProperties.getProperty('templateId') || null;
const templateName = scriptProperties.getProperty('templateName') || null;

// Create PandaDoc menu
function onOpen() {
    let ui = SpreadsheetApp.getUi();
    if (templateId && APIKey && !templateName) getTemplateName(templateId);

    // If there is no API key, render just the API key prompt option
    if (!APIKey) {
        ui.createMenu('PandaDoc')
            .addItem('Add API key (No API key saved, enter API key and refresh page to access other options)', 'showKeyPrompt')
            .addToUi();
    } else if (APIKey && !sandboxAPIKey) {
        ui.createMenu('PandaDoc')
            .addItem('Create 1 Document (disabled, please add sandbox API key to use)', 'disabled')
            .addItem('Create all Documents', 'createAllDocuments')
            .addItem('Send all Documents', 'sendAllDocs')
            .addSeparator()
            .addItem('Get all content library items', 'fetchCLI')
            .addSeparator()
            .addItem('Update API key', 'showKeyPrompt')
            .addItem('Add sandbox API key', 'showSandKeyPrompt')
            .addItem(`Set Template (${setTemplateMenuText()})`, 'showTemplatePrompt')
            .addToUi();
    } else {
        ui.createMenu('PandaDoc')
            .addItem('Create 1 Document (sandbox testing)', 'createOneDocument')
            .addItem('Create all Documents', 'createAllDocuments')
            .addItem('Send all Documents', 'sendAllDocs')
            .addSeparator()
            .addItem('Get all content library items', 'fetchCLI')
            .addSeparator()
            .addItem('Update API key', 'showKeyPrompt')
            .addItem('Update sandbox API key', 'showSandKeyPrompt')
            .addItem(`Set Template (${setTemplateMenuText()})`, 'showTemplatePrompt')
            .addToUi();
    }
}

const disabled = () => {
    let ui = SpreadsheetApp.getUi(); // Same variations.
    let result = ui.alert(
        'disabled',
        'This feature is disabled because the sandbox API key is missing.',
        ui.ButtonSet.OK
    );
    return
}

// Get the display text for the template menu option
const setTemplateMenuText = () => {
    return templateName || 'no template set'
}

// Get the template name for display purposes
const getTemplateName = async (id) => {
    let template = await UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/templates/${id}/details`, {
        method: 'GET',
        headers: {
            'Authorization': `Api-key ${APIKey}`
        },
    })
    let responseCode = template.getResponseCode();
    if (responseCode != 200) {
        Logger.log("Error getting docs: " + responseCode + " body: " + responseText);
        throw "Error getting docs: " + responseCode + " body: " + responseText;
    }
    template = JSON.parse(template);
    let name = template.name;
    scriptProperties.setProperty('templateName', name);
    return name
}

// Retreieve all template details
const getTemplateDetails = async (id) => {
    let template = await UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/templates/${id}/details`, {
        method: 'GET',
        headers: {
            'Authorization': `Api-key ${APIKey}`
        },
    })
    let responseCode = template.getResponseCode();
    if (responseCode != 200) {
        Logger.log("Error getting docs: " + responseCode + " body: " + responseText);
        throw "Error getting docs: " + responseCode + " body: " + responseText;
    }
    template = JSON.parse(template);
    let name = template.name;
    scriptProperties.setProperty('templateName', name);
    return template
}

// Set API key
const showKeyPrompt = () => {
    let apiKey = scriptProperties.getProperty('apiKey')
    let ui = SpreadsheetApp.getUi(); // Same variations.
    let result = ui.prompt(
        'apiKey ',
        'Current Key: ' + apiKey,
        ui.ButtonSet.OK_CANCEL);

    // Process the user's response.
    let button = result.getSelectedButton();
    let key = result.getResponseText();
    if (button == ui.Button.OK) {
        scriptProperties.setProperty('apiKey', key);

        // User clicked "OK".
        ui.alert('Your key ' + key + ' was saved. Refresh page to update PandaDoc menu.');
    }
}

// Set sandbox API key
const showSandKeyPrompt = () => {
    let sandboxApiKey = scriptProperties.getProperty('sandboxApiKey')
    let ui = SpreadsheetApp.getUi(); // Same variations.
    let result = ui.prompt(
        'Set sandbox API key',
        'Current Key: ' + sandboxAPIKey,
        ui.ButtonSet.OK_CANCEL);

    // Process the user's response.
    let button = result.getSelectedButton();
    let key = result.getResponseText();
    if (button == ui.Button.OK) {
        scriptProperties.setProperty('sandboxApiKey', key);
        // User clicked "OK".
        ui.alert('Your key ' + key + ' was saved. Refresh page to update PandaDoc menu.');
    }
}

// Set template
const showTemplatePrompt = async () => {
    let ui = SpreadsheetApp.getUi(); // Same variations.

    let response = ui.alert('WARNING: This action will clear all data from the document sheet, are you sure you want to continue?', ui.ButtonSet.YES_NO);

    if (response == ui.Button.YES) {
        let result = ui.prompt(
            'Set template ID ',
            'Current template: ' + setTemplateMenuText(),
            ui.ButtonSet.OK_CANCEL);
        // Process the user's response.
        let button = result.getSelectedButton();
        let id = result.getResponseText();
        if (button == ui.Button.OK) {

            scriptProperties.setProperty('templateId', id);


            let template = await getTemplateDetails(id);

            updateSheet(template)

            // User clicked "OK".
            ui.alert(`Your id ${id} (${template.name}) was saved.`);
        }
    }
}

// Delete all content in a sheet
const clearSheet = (sheetName) => {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    let data = sheet.getDataRange().getValues();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setDataValidation(null);
    let headers = data[0];
    sheet.clearContents();
}

// Get the column coordinates for the data validation cell
const getSheetColumn = (index) => {
    let alphabet = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
    let column = "";
    if (index > alphabet.length - 1) {
        column += alphabet[Math.floor(index / alphabet.length) - 1];
        column += alphabet[index % 26 - 1];
    } else {
        column = alphabet[index];
    }
    return column
}

// Set headers, first row of sheet
const updateSheet = async (template) => {
    let headers = ["Document_URL", "Document_uuid", "Document_status", "email_subject", "email message"];
    let firstRow = ["", "", "", "", ""];

    // Delete all content in the documents sheet, set let sheet as the documents sheet
    clearSheet("Documents");
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Documents");
    let CLISheeet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Content Libary items");

    // create headers from roles
    template.roles.forEach(role => {
        headers.push(`role_${role.name}_first_name`)
        firstRow.push("")
        headers.push(`role_${role.name}_last_name`)
        firstRow.push("")
        headers.push(`role_${role.name}_email`)
        firstRow.push("")
    })

    // create headers from content place holder 
    template.content_placeholders.forEach(content => {
        headers.push(`CPH_${content.block_id}`)
        let sheetIndex = headers.length - 1;
        let cell = sheet.getRange(`${getSheetColumn(sheetIndex)}2:${getSheetColumn(sheetIndex)}100`);
        let range = CLISheeet.getRange('B2:B200');
        let rule = SpreadsheetApp.newDataValidation().requireValueInRange(range)
            .build();
        cell.setValue("")
        cell.setDataValidation(rule);
    })

    // create headers from variables 
    template.tokens.forEach(token => headers.push(`VAR_${token.name}`))

    // put headers on sheet
    sheet.appendRow(headers);
    // adjust all column widths
    sheet.autoResizeColumns(1, headers.length);
    return
}

// Get all content library items 
const fetchCLI = async () => {
    clearSheet("Content Libary items");
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Content Libary items")
    let data = sheet.getDataRange().getValues();
    let isFetching = true;
    let page = 1;

    do {
        let items = await UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/content-library-items?count=100&page=${page}`, {
            method: 'GET',
            headers: {
                'Authorization': `Api-key ${APIKey}`
            },
        });
        let responseCode = items.getResponseCode();
        if (responseCode != 200) {
            Logger.log("Error getting docs: " + responseCode + " body: " + responseText);
            throw "Error getting docs: " + responseCode + " body: " + responseText;
        }
        items = JSON.parse(items);

        items = items.results.map(item => ({
            id: item.id,
            name: item.name,
            date_created: item.date_created
        }))
        isFetching = items.length === 100;
        let sheetHeaders = Object.keys(items[0])
        items = items.map(item => Object.values(item))
        if (page === 1) items.splice(0, 0, sheetHeaders);

        items.forEach(item => {
            sheet.appendRow(item);
        })
        page++;

    } while (isFetching)
}

const documentName = (sheet, headers, i) => {
        let indexFN = headers.indexOf('role_Plan Recipient_first_name') + 1; 
        let indexLN = headers.indexOf('role_Plan Recipient_last_name') + 1;
        let first_name = sheet.getRange(i + 2, indexFN).getValue();
        let last_name = sheet.getRange(i + 2, indexLN).getValue();
    return {
      first_name,
      last_name
    }
}

const createAllDocuments = async () => {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Documents");
    let range = sheet.getDataRange().getValues();
    let headers = range.splice(0, 1)[0];


    // Iterate through rows
    range.forEach(async (doc, i) => {
        let statusIndex = headers.indexOf('Document_status');
        if (doc[statusIndex]) return;

        let names = documentName(sheet, headers, i)
        console.log(names.first_name);

        let payload = {
            template_uuid: scriptProperties.getProperty('templateId'),
            name: templateName + " " + names.first_name + " " + names.last_name,
        };

        doc.forEach(async (value, j) => {
            payload = assignValues(headers[j], value, payload);
        })
        
        let res = await UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json;charset=UTF-8',
                'Authorization': `API-key ${APIKey}`
            },
            payload: JSON.stringify(payload)
        })
        let responseCode = res.getResponseCode();
        if (responseCode != 200 && responseCode != 201) {
            Logger.log("Error creating doc: " + responseCode + " row: " + (i + 2));
            throw "Error creating doc: " + responseCode + " row: " + (i + 2);
        }
        res = JSON.parse(res);
        let docRange = sheet.getRange(`A${i + 2}:C${i + 2}`);
        docRange.setValues([
            [`https://app.pandadoc.com/a/#/documents/${res.id}`, res.id, res.status]
        ])
    });
}

const createOneDocument = async () => {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Documents");
    let range = sheet.getDataRange().getValues();
    let headers = range.splice(0, 1)[0];
    let CLISheeet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Content Libary items");


    // Iterate through rows
    for (let i = 0; i < range.length; i++) {
        let doc = range[i];
        let payload = {
            template_uuid: scriptProperties.getProperty('templateId'),
            name: sheet.getRange(F[i]).getValue(), //scriptProperties.getProperty('templateName') + 
        };

        let statusIndex = headers.indexOf('Document_status');
        // if there is already a doc status, don't continue with this row
        if (!doc[statusIndex]) {
            doc.forEach(async (value, j) => {
                payload = assignValues(headers[j], value, payload);
            })

            let res = await UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json;charset=UTF-8',
                    'Authorization': `API-key ${sandboxAPIKey}`
                },
                payload: JSON.stringify(payload)
            })

            let responseCode = res.getResponseCode();
            if (responseCode != 200 && responseCode != 201) {
                Logger.log("Error getting docs: " + responseCode + " row: " + (i + 2));
                throw "Error getting docs: " + responseCode + " row: " + (i + 2);
            }
            res = JSON.parse(res);
            let docRange = sheet.getRange(`A${i + 2}:C${i + 2}`);
            docRange.setValues([
                [`https://app.pandadoc.com/a/#/documents/${res.id}`, res.id, res.status]
            ])
            break;
        }

    }
}


const assignValues = (key, value, payload) => {
    let result;

    if (key.includes('role_')) {
        result = roleValue(key, value, payload)
    } else if (key.includes('CPH_')) {
        if ('value')
            result = contentValue(key, value, payload)
    } else if (key.includes('VAR_')) {
        result = variableValue(key, value, payload)
    } else {
        result = payload
    }

    return result
}

const roleValue = (key = "role_test_first_name", value, payload) => {
    let myRegexp = /role_(.*)/;
    let match = myRegexp.exec(key);
    let re_key = match[1];
    // Need to get the role name from the header, as well as as the parameter we are passing to the role 
    let keyList = re_key.split('_');
    let roleName = keyList.splice(0, 1)[0];
    re_key = keyList.join('_');
    if (!('recipients' in payload)) payload.recipients = [{
        role: roleName
    }];
    let recipients = payload.recipients.filter(contact => Object.keys(contact).length < 4);
    let contact;
    if (!recipients.length) {
        payload.recipients.push({
            role: roleName
        })
        contact = payload.recipients[payload.recipients.length - 1]
    } else {
        contact = recipients[0]
    }
    contact[re_key] = value;
    return payload
}

const contentValue = (key, value, payload) => {
    let myRegexp = /CPH_(.*)/;
    let match = myRegexp.exec(key);
    let re_key = match[1];
    // Variables for finding the content library item ID 
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Content Libary items');
    let data = sheet.getDataRange().getValues();
    let headers = data.splice(0, 1)[0];
    let idIndex = headers.indexOf('id');
    let nameIndex = headers.indexOf('name');
    let itemId;

    // Find the id for the content library item based on the name provided in the cell
    for (let i = 0; i < data.length - 1; i++) {
        if (data[i][nameIndex] === value) {
            itemId = data[i][idIndex];
            break;
        }
        // If we've reached the end of the sheet and there is no name that matches the value in the cell, something has gone wrong, this shouldn't really be possible unless the user manually entered the value instead of choosing from the drop down
        if (i === data.length - 1 && !itemId) throw `Couldn't find ID for content placeholder in the 'Content Libary items' sheet`;
    }

    if (!('content_placeholders' in payload)) payload.content_placeholders = [];

    let content = {
        block_id: re_key,
        content_library_items: [{
            id: itemId
        }]
    };

    payload.content_placeholders.push(content)
    return payload

}

const variableValue = (key, value, payload) => {
    let myRegexp = /VAR_(.*)/;
    let match = myRegexp.exec(key);
    let re_key = match[1];

    if (!('tokens' in payload)) payload.tokens = [];

    let token = {
        name: re_key,
        value
    };
    payload.tokens.push(token);
    return payload
}

const sendAllDocs = async () => {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Documents");
    let len = SpreadsheetApp.getActiveSheet().getLastRow();

    for (let i = 2; i <= len; i++) {
        let id = sheet.getRange(`B${i}`).getValue();
        let message = sheet.getRange(`E${i}`).getValue();
        let subject = sheet.getRange(`D${i}`).getValue();

        let statusRes = await UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents/${id}`, {
            method: 'GET',
            headers: {
                'Authorization': `API-key ${APIKey}`
            },
        })
        statusRes = JSON.parse(statusRes);
        let status = statusRes.status;

        if (status === 'document.draft') {
            let res = await UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents/${id}/send`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json;charset=UTF-8',
                    'Authorization': `API-key ${APIKey}`,
                },
                'payload': JSON.stringify({
                    message: message,
                    subject: subject
                })
            })
            let result = await JSON.parse(res);
            console.log(result);
            let docRange = sheet.getRange(`C${i}`);
            docRange.setValues([
                [result.status]
            ]);
            Utilities.sleep(1000);
        }
    }
}