var request = require('request-promise');
const Promise = require('promise');
const config = require('config');
const { google } = require('googleapis');
const fs = require('fs');
const readline = require('readline');
const luxon = require('luxon');
const util = require('util');

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_PATH = 'config/token.json';

const falinux = {lon: "126.99024683", lat: "37.40150134" }
const hjauto = { lon: "126.88114364", lat: "37.47296332" }
const home = { lon: "126.82806535", lat: "37.46551880" }


let columns = [];
for (c = 'A'.charCodeAt(0); c <= 'Z'.charCodeAt(0); c++) {
	columns.push(String.fromCharCode(c));
}
for (prefix = 'A'.charCodeAt(0); prefix <= 'Z'.charCodeAt(0); prefix++) {
    for (c = 'A'.charCodeAt(0); c <= 'Z'.charCodeAt(0); c++) {
        columns.push(String.fromCharCode(prefix) + String.fromCharCode(c));
    }
}
let weekdays = ['월', '화', '수', '목', '금', '토', '일'];

var authorize = function(oAuth2Client) {
    return new Promise((resolve, reject) => {
        // Check if we have previously stored a token.
        fs.readFile(TOKEN_PATH, (err, token) => {
            if (err) {
                const authUrl = oAuth2Client.generateAuthUrl({
                    access_type: 'offline',
                    scope: SCOPES,
                });
                console.log('Authorize this app by visiting this url:', authUrl);
                const rl = readline.createInterface({
                    input: process.stdin,
                    output: process.stdout,
                });
                rl.question('Enter the code from that page here: ', (code) => {
                    rl.close();
                    oAuth2Client.getToken(code, (err, token) => {
                        if (err) {
                            console.error('Error retrieving access token', err);
                            reject(err);
                            return;
                        }
                        oAuth2Client.setCredentials(token);
                        // Store the token to disk for later program executions
                        fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                            if (err) {
                                console.error(err);
                                reject(err);
                                return;
                            }
                            console.log('Token stored to', TOKEN_PATH);
                        });
            
                        resolve(oAuth2Client);
                    });
                });
            }
            oAuth2Client.setCredentials(JSON.parse(token));
            resolve(oAuth2Client);
        });
    });
}

var tmap_trace = function(start, end) {
    var requestOption = {
        uri: 'https://api2.sktelecom.com/tmap/routes?version=1',
        method: 'POST',
        form: {
            startX: start['lon'],
            startY: start['lat'],
            endX: end['lon'],
            endY: end['lat'],
            reqCoordType: "WGS84GEO",
            resCoordType: "EPSG3857",
            searchOption: "0",
            trafficInfo: "Y"
        },
        headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'ko,en-US;q=0.8,en;q=0.6',
            'appKey': config.get('tmap').appKey
        },
        jar: true,
        gzip: true,
        encoding: null
    };

    return request(requestOption);
}

const develop_spreadsheets = [
    {id: '1_HcGNs1XylAaEKu1NwIRGaPJn0wS42-v6OiVguhUO9M', path_list: [{name: "출근", start: home, end: falinux, time: 0}, {name: "퇴근", start: falinux, end: home, time: 0}]}
]

const service_spreadsheets = [
    {id: '1_HcGNs1XylAaEKu1NwIRGaPJn0wS42-v6OiVguhUO9M', path_list: [{name: "출근", start: home, end: hjauto, time: 0}, {name: "퇴근", start: hjauto, end: home, time: 0}]},
    //{id: '1NYHVggzwYViUA7dE_i2sUKm5oZmMJrW-U1Drwb_ZNnc', path_list: [{name: "출근", start: home, end: falinux, time: 0}, {name: "퇴근", start: falinux, end: home, time: 0}]}
]

exports.handler = function (event, context, callback) {
    var start_date = luxon.DateTime.fromISO("2019-06-30T15:00:00Z").setZone('Asia/Seoul');
    var today = luxon.DateTime.local().setZone('Asia/Seoul');
    var row = Math.floor((today - start_date)/24/60/60/1000) + 2;
    
    const { client_secret, client_id, redirect_uris } = config.get('installed');
    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

    var promiseList = [];
    authorize(oAuth2Client).then((auth) => {
        const service = google.sheets({ version: 'v4', auth });

        service_spreadsheets.forEach((sheet) => {
            sheet.path_list.forEach((path) => {
                promiseList.push(new Promise((resolve, reject) => {
                    service.spreadsheets.values.clear({
                        spreadsheetId: sheet.id,
                        range: util.format('%s!A%d:B%d', path.name, row, row)
                    }, (err, res) => {
                        /*
                        if (err) {
                            console.error('The API returned an error: ' + err);
                            reject(err);
                            return;
                        }
                        */
    
                        service.spreadsheets.values.append({
                            spreadsheetId: sheet.id,
                            range: util.format('%s!A%d:B%d', path.name, row, row),
                            valueInputOption: 'USER_ENTERED',
                            resource: { values: [[today.toFormat("yyyy-MM-dd"), weekdays[today.weekday - 1]]] }
                        }, (err, res) => {
                            if (err) {
                                console.error('The API returned an error: ' + err);
                                reject(err);
                                return;
                            }
                            console.log(path.name, "날짜/요일 입력 완료");
                            resolve();
                        });
                    });
                }));

                promiseList.push(tmap_trace(path.start, path.end).then((html) => {
                    return new Promise((resolve, reject) => {
                        obj = JSON.parse(html);
                        path.time = obj["features"][0]["properties"]["totalTime"];
                        console.log(path.name, "경로 예상 시간 측정 완료", path.time);

                        service.spreadsheets.values.update({
                            spreadsheetId: sheet.id,
                            range: util.format('%s!%s%d', path.name, columns[Math.floor(today.hour * 6 + today.minute / 10) + 2], row),
                            valueInputOption: 'USER_ENTERED',
                            resource: { values: [[path.time/(24.0*60*60)]] }
                        }, (err, res) => {
                            if (err) {
                                console.error('The API returned an error: ' + err);
                                reject(err);
                                return;
                            }
                            console.log(path.name, "시간 입력", res.data.updatedRange, ":", res.statusText);
                            resolve();
                        });
                    });
                }));
            });
        });
    }).then(() => {
        Promise.all(promiseList).then(() => {
            console.log("Done");
            
            if (callback) {
                callback(null, 'Success');
            }
        })
    });

};
