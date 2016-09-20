'use strict';

var fs = require('fs');
var readline = require('readline');
var google = require('googleapis');
var googleAuth = require('google-auth-library');
const path = require('path');

const SPREADSHEET_ID = "1spn_v2h6Q_nr9nnqQjjV-JNrx04u3iszSkBZzqfeNzU";
const SPREADSHEET_ID_2 = "1idKkmKO_nLG2ly9ifQ25qVVo37-oJWTzllRNqQ6iTck";

let lastRow = 3;

// If modifying these scopes, delete your previously saved credentials
// at ~/.credentials/sheets.googleapis.com-nodejs-quickstart.json

var SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'];
var TOKEN_DIR = path.join(__dirname, '/.credentials/');
var TOKEN_PATH = TOKEN_DIR + 'token.json';

// Load client secrets from a local file.
fs.readFile('client_secret.json', (err, content) => {
    if (err) {
        console.log('Error loading client secret file: ' + err);
        return;
    }
    // Authorize a client with the loaded credentials, then call the Google Sheets API.
    authorize(JSON.parse(content), appendValues);
});

const authorize = (credentials, callback) => {
    var clientSecret = credentials.installed.client_secret;
    var clientId = credentials.installed.client_id;
    var redirectUrl = credentials.installed.redirect_uris[0];
    var auth = new googleAuth();
    var oauth2Client = new auth.OAuth2(clientId, clientSecret, redirectUrl);

    // Check if we have previously stored a token.
    fs.readFile(TOKEN_PATH, (err, token) => {
        if (err) {
            getNewToken(oauth2Client, callback);
        } else {
            oauth2Client.credentials = JSON.parse(token);
            callback(oauth2Client);
        }
    });
};

const getNewToken = (oauth2Client, callback) => {
    var authUrl = oauth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES
    });
    console.log('Authorize this app by visiting this url: ', authUrl);
    var rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });
    rl.question('Enter the code from that page here: ', (code) => {
        rl.close();
        oauth2Client.getToken(code, (err, token) => {
            if (err) {
                console.log('Error while trying to retrieve access token', err);
                return;
            }
            oauth2Client.credentials = token;
            storeToken(token);
            callback(oauth2Client);
        });
    });
};

const storeToken = (token) => {
    try {
        fs.mkdirSync(TOKEN_DIR);
    } catch (err) {
        if (err.code != 'EEXIST') {
            throw err;
        }
    }
    fs.writeFile(TOKEN_PATH, JSON.stringify(token));
    console.log('Token stored to ' + TOKEN_PATH);
};

/**
 * Print the names and majors of students in a sample spreadsheet:
 * https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 */
const insertAtLastRow = (auth) => {
    const sheets = google.sheets('v4');
    const options = {
        auth: auth,
        spreadsheetId: SPREADSHEET_ID,
        range: `Sheet1!A${lastRow + 1}`,
        valueInputOption: "RAW",
        resource: {
            majorDimension: "ROWS",
            values: [
                ['Works!!']
            ]
        }
    };
    sheets.spreadsheets.values.update(options, (err, res) => {
        if (err) {
            console.log(`The API returned an error: ${err}`);
            return;
        }
        console.log(JSON.stringify(res, null, 2));
    });
};

const batchUpdate = (auth) => {
    const sheets = google.sheets('v4');
    let params = {
        auth,
        spreadsheetId: SPREADSHEET_ID_2,
        resource: {
            requests: [{
                "appendDimension": {
                    "sheetId": 0,
                    "dimension": "ROWS",
                    "length": 3
                }
            }]
        }
    };
    sheets.spreadsheets.batchUpdate(params, (err, response) => {
        console.log("Response: ", response);
    });
}

const appendValues = (auth) => {
    const sheets = google.sheets('v4');
    let params = {
        auth,
        valueInputOption: "USER_ENTERED",
        spreadsheetId: SPREADSHEET_ID_2,
        range: "Sheet1!A1:E1",
        resource: {
            range: "Sheet1!A1:E1",
            majorDimension: "ROWS",
            values: [
                ["4 Door", "$15", "2", "3/15/2016"]
            ]
        }
    };
    sheets.spreadsheets.values.append(params, (err, response) => {
        console.log("Response: ", err, response);
    });
}

const readSheet = (auth) => {
    const sheets = google.sheets('v4');
    const options = {
        auth: auth,
        spreadsheetId: SPREADSHEET_ID,
        range: 'Sheet1'
    };

    sheets.spreadsheets.values.get(options, (err, res) => {
        if (err) {
            console.log(`The API returned an error: ${err}`);
            return;
        }
        console.log(JSON.stringify(res, null, 2));
        lastRow = res.values.length;
        console.log(res.values.length);
    });
};

const getSpreadsheet = (auth) => {
    const sheets = google.sheets('v4');
    const options = {
        spreadsheetId: SPREADSHEET_ID_2,
        auth: auth
    };
    sheets.spreadsheets.get(options, (err, res) => {
        if (err) {
            console.log(`The API returned an error: ${err}`);
            return;
        }
        console.log(JSON.stringify(res, null, 2));
    });
}

const createSheet = (auth) => {
    const sheets = google.sheets('v4');
    const options = {
        auth: auth,
        resource: {
            properties: {
                title: 'My Api Test'
            }
        }
    };
    sheets.spreadsheets.create(options, (err, res) => {
        if (err) {
            console.log(`The API returned an error: ${err}`);
            return;
        }
        console.log(JSON.stringify(res, null, 2));
    });
};

const getSheet = (auth) => {
    const sheets = google.sheets('v4');
    const options = {
        auth: auth,
        spreadsheetId: SPREADSHEET_ID_2
    };
    sheets.spreadsheets.get(options, (err, res) => {
        if (err) {
            console.log(`The API returned an error: ${err}`);
            return;
        }
        console.log(JSON.stringify(res, null, 2));
    });
};

const listFiles = (auth) => {
    let service = google.drive('v3');
    let options = {
        auth: auth,
        pageSize: 50,
        fields: "nextPageToken, files(id, name)"
    };
    service.files.list(options, (err, res) => {
        if (err) {
            console.log(`The API returned an error: ${err}`);
            return;
        }
        let files = res.files;
        if (files.length == 0) {
            console.log(`No files found`);
            return;
        }
        files.forEach(file => {
            console.log(`${file.name}\n${file.id} \n\n`);
        });
    });
};

const listMajors = (auth) => {
    var sheets = google.sheets('v4');
    sheets.spreadsheets.values.get({
        auth: auth,
        spreadsheetId: '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
        range: 'Class Data!A2:E',
    }, (err, response) => {
        if (err) {
            console.log('The API returned an error: ' + err);
            return;
        }
        var rows = response.values;
        if (rows.length == 0) {
            console.log('No data found.');
        } else {
            console.log('Name, Major:');
            for (var i = 0; i < rows.length; i++) {
                var row = rows[i];
                // Print columns A and E, which correspond to indices 0 and 4.
                console.log('%s, %s', row[0], row[4]);
            }
        }
    });
};
