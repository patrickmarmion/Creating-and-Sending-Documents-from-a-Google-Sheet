const scriptProperties = PropertiesService.getScriptProperties();
const APIKey = scriptProperties.getProperty('apiKey') || null;
const templateId = scriptProperties.getProperty('templateId') || null;
const templateName = scriptProperties.getProperty('templateName') || null;

// Create PandaDoc menu
const onOpen = () => {
    let ui = SpreadsheetApp.getUi();

    // If there is no API key, render just the API key prompt option
    if (!APIKey) {
        ui.createMenu('PandaDoc')
            .addItem('Add API key (No API key saved, enter API key and refresh page to access other options)', 'showKeyPrompt')
            .addToUi();
    } else {
        ui.createMenu('PandaDoc')
            .addItem('Create Document', 'createDocument')
            .addToUi();
    }
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

const createdDocValues = (res, sheet, i) => {
    let values = [
        [`https://app.pandadoc.com/a/#/documents/${res.id}`, res.id, res.status]
    ];
    let range = sheet.getRange(`A${i}:C${i}`);
    range.setValues(values);
    return
}

const createDocument = async () => {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let i = sheet.getLastRow();

    let payload = {};
    //Here the Document Name is: Template Name + Recipient First Name + Receipient Last Name
    payload.name = templateName + sheet.getRange(`F${i}`).getValue() + sheet.getRange(`G${i}`).getValue();
    payload.template_uuid = templateId;

    payload.recipients = [{
        "email": sheet.getRange(`H${i}`).getValue(),
        "first_name": sheet.getRange(`F${i}`).getValue(),
        "last_name": sheet.getRange(`G${i}`).getValue(),
        "role": "Client"
    }];

    payload.tokens = [{
            "name": "Example_Token",
            "value": sheet.getRange(`K${i}`).getValue()
        },
        {
            "name": "Another_Example",
            "value": sheet.getRange(`L${i}`).getValue()
        }
    ];

    let createOptions = {
        'method': 'post',
        'headers': {
            'Authorization': `API-Key ${APIKey}`,
            'Content-Type': 'application/json;charset=UTF-8'
        },
        'payload': JSON.stringify(payload)
    };

    let response = UrlFetchApp.fetch("https://api.pandadoc.com/public/v1/documents", createOptions);
    let responseCode = response.getResponseCode();
    if (responseCode != 200 && responseCode != 201) {
        Logger.log("Error creating doc: " + responseCode);
        throw "Error creating doc: " + responseCode;
    }
    let responseJson = JSON.parse(response);
    console.log(responseJson);

    await createdDocValues(responseJson, sheet, i)
}