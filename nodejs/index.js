var request = require('request');
const config = require('config');
const { google } = require('googleapis');
const fs = require('fs');
const readline = require('readline');
const luxon = require('luxon');
const util = require('util');
const async = require('async');
const AWS = require('aws-sdk');

AWS.config.update({
    region: 'ap-northeast-2',
    endpoint: "http://dynamodb.ap-northeast-2.amazonaws.com"
});

//const dynamodb = new AWS.DynamoDB();
const docClient = new AWS.DynamoDB.DocumentClient();

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_PATH = 'config/token.json';

const home = { lon: "126.82806535", lat: "37.46551880" };
const company1 = {lon: "127.090294", lat: "37.391982"};

const develop_spreadsheets = [
    { id: '1BOozLi2KsCemNhMETZaZVNPfQK3IgpDMVw9QQy0P5wI', name: "출근", start: home, end: company1, time: 0 },
    { id: '1BOozLi2KsCemNhMETZaZVNPfQK3IgpDMVw9QQy0P5wI', name: "퇴근", start: company1, end: home, time: 0 },
];

const service_spreadsheets = [
    { id: '1_HcGNs1XylAaEKu1NwIRGaPJn0wS42-v6OiVguhUO9M', name: "출근", start: home, end: company1, time: 0, since: "2025-02-10T15:00:00Z", appKey: config.get('tmap').appKeyWORK },
    { id: '1_HcGNs1XylAaEKu1NwIRGaPJn0wS42-v6OiVguhUO9M', name: "퇴근", start: company1, end: home, time: 0, since: "2025-02-10T15:00:00Z", appKey: config.get('tmap').appKeyHOME }
];

const target_sheets = service_spreadsheets;
const viewer_sheets = {data: '1_HcGNs1XylAaEKu1NwIRGaPJn0wS42-v6OiVguhUO9M', viewer: '1aVBgwXfhn4lv13cvjJ52rz7GJV92iXojSJt35pxmZoQ'}

const time_period = 5;

let columns = [];
for (var c = 'A'.charCodeAt(0); c <= 'Z'.charCodeAt(0); c++) {
    columns.push(String.fromCharCode(c));
}
for (var prefix = 'A'.charCodeAt(0); prefix <= 'Z'.charCodeAt(0); prefix++) {
    for (var c = 'A'.charCodeAt(0); c <= 'Z'.charCodeAt(0); c++) {
        columns.push(String.fromCharCode(prefix) + String.fromCharCode(c));
    }
}
let weekdays = ['월', '화', '수', '목', '금', '토', '일'];

var buildStatisticsFormula = function (sheet, index, callback) {
    sheet.statistics_range = util.format("%s요약!%s2:%s", sheet.name, columns[1], columns[7]);
    sheet.statistics_values = [];

    for (var row = 2; row < (24 * 60 / time_period) + 2; row++) {
        let values = [];
        for (var col = 1; col < 8; col++) {
            var f = `=IFERROR(Floor(TRIMMEAN(FILTER('${sheet.name}'!$${columns[row]}$2:$${columns[row]}, '${sheet.name}'!$B$2:$B = ${columns[col]}$1), 0.25), (1 * 60)/(24*60*60)), \"\")`;
            values.push(f);
        }
        sheet.statistics_values.push(values);
    }

    callback(null, sheet, index);
};

var clearStatisticsFormula = function (sheet, index, callback) {
    sheet.service.spreadsheets.values.clear({
        spreadsheetId: sheet.id,
        range: sheet.statistics_range
    }, (err, res) => {
        callback(null, sheet, index);
    });
};

var updateStatisticsFormula = function (sheet, index, callback) {
    console.log(sheet.id, sheet.name, "통계 테이블 입력 요청");
    sheet.service.spreadsheets.values.update({
        spreadsheetId: sheet.id,
        range: sheet.statistics_range,
        valueInputOption: 'USER_ENTERED',
        resource: { values: sheet.statistics_values }
    }, (err, res) => {
        if (!err) {
            console.log(sheet.id, sheet.name, "통계 테이블 입력 완료");
        }
        callback(err, sheet, index);
    });
};

var buildSummaryFormula = function (sheet, index, callback) {
    if (sheet.name === "출근") {
        sheet.summary_range = util.format("요약!%s1:%s", columns[1], columns[3]);
        sheet.summary_values = [[`평균 ${sheet.name}`, `요일 ${sheet.name}`, `오늘 ${sheet.name}`]];
    }
    if (sheet.name === "퇴근") {
        sheet.summary_range = util.format("요약!%s1:%s", columns[4], columns[6]);
        sheet.summary_values = [[`평균 ${sheet.name}`, `요일 ${sheet.name}`, `오늘 ${sheet.name}`]];
    }

    for (var row = 2; row < (24 * 60 / time_period) + 2; row++) {
        var summary_row = [];

        if (sheet.name === "출근") {
            summary_row.push(util.format("=IFERROR(IF(MOD($A%d*24, 24) < 12, Floor(TRIMMEAN(query('%s'!$%s$2:$%s, \"select \`%s\` where B <> '토' and B <> '일' order by A desc limit 180\"), 0.25), (1 * 60)/(24*60*60)), \"\"), \"\")", row, sheet.name, "A", "KD", columns[row]));
            summary_row.push(util.format("=IFERROR(IF(MOD($A%d*24, 24) < 12, Floor(TRIMMEAN(query('%s'!$%s$2:$%s, \"select \`%s\` where B = '\"&TEXT(TODAY(), \"ddd\")&\"' order by A desc limit 180\"), 0.25), (1 * 60)/(24*60*60)), \"\"), \"\")", row, sheet.name, "A", "KD", columns[row]));
            summary_row.push(util.format("=IFERROR(IF(MOD($A%d*24, 24) < 12, Floor(AVERAGEIFS('%s'!$%s$2:$%s, '%s'!$A$2:$A, TODAY()), (1 * 60)/(24*60*60)), \"\"), \"\")", row, sheet.name, columns[row], columns[row], sheet.name));
            sheet.summary_values.push(summary_row);
        }

        if (sheet.name === "퇴근") {
            summary_row.push(util.format("=IFERROR(IF(MOD($A%d*24, 24) >= 12, Floor(TRIMMEAN(query('%s'!$%s$2:$%s, \"select \`%s\` where B <> '토' and B <> '일' order by A desc limit 180\"), 0.25), (1 * 60)/(24*60*60)), \"\"), \"\")", row, sheet.name, "A", "KD", columns[row]));
            summary_row.push(util.format("=IFERROR(IF(MOD($A%d*24, 24) >= 12, Floor(TRIMMEAN(query('%s'!$%s$2:$%s, \"select \`%s\` where B = '\"&TEXT(TODAY(), \"ddd\")&\"' order by A desc limit 180\"), 0.25), (1 * 60)/(24*60*60)), \"\"), \"\")", row, sheet.name, "A", "KD", columns[row]));
            summary_row.push(util.format("=IFERROR(IF(MOD($A%d*24, 24) >= 12, Floor(AVERAGEIFS('%s'!$%s$2:$%s, '%s'!$A$2:$A, TODAY()), (1 * 60)/(24*60*60)), \"\"), \"\")", row, sheet.name, columns[row], columns[row], sheet.name));
            sheet.summary_values.push(summary_row);
        }
    }

    callback(null, sheet, index);
};

var clearSummaryFormula = function (sheet, index, callback) {
    sheet.service.spreadsheets.values.clear({
        spreadsheetId: sheet.id,
        range: sheet.summary_range
    }, (err, res) => {
        callback(null, sheet, index);
    });
};

var updateSummaryFormula = function (sheet, index, callback) {
    console.log(sheet.id, sheet.name, "요약 테이블 입력 요청");
    sheet.service.spreadsheets.values.update({
        spreadsheetId: sheet.id,
        range: sheet.summary_range,
        valueInputOption: 'USER_ENTERED',
        resource: { values: sheet.summary_values }
    }, (err, res) => {
        if (!err) {
            console.log(sheet.id, sheet.name, "요약 테이블 입력 완료");
        }
        callback(err, sheet, index);
    });
};

exports.handle_formula = function (event, context, callback) {
    async.waterfall([
        function (callback) {
            const { client_secret, client_id, redirect_uris } = config.get('installed');
            callback(null, {
                oAuth2Client: new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]),
                automated: false,
                tokenReady: false,
                sheets: target_sheets
            });
        },
        authorize,
        makeToken,
        function (args, callback) {
            const service = google.sheets({ version: 'v4', auth: args.oAuth2Client });

            async.eachOf(args.sheets, function (sheet, index, callback) {
                sheet.service = service;

                async.waterfall([
                    function (callback) {
                        callback(null, sheet, index);
                    },
                    buildSummaryFormula,
                    clearSummaryFormula,
                    updateSummaryFormula,
                    buildStatisticsFormula,
                    clearStatisticsFormula,
                    updateStatisticsFormula,
                ], function (err) {
                    if (err) {
                        console.log(err);
                    }
                    callback(err);
                });
            }, function (err) {
                if (err) {
                    console.log(err);
                }
                callback(err, args);
            });
        }
    ], function (err) {
        if (err) {
            console.log(err);
        }
    });
};

exports.handle_location = function (event, context, callback) {
    async.waterfall([
        foundLocation,
    ], function (err) {
        if (err) {
            console.log(err);
        }
    });
};

var authorize = function (args, callback) {
    // Check if we have previously stored a token.
    fs.readFile(TOKEN_PATH, (err, token) => {
        if (err) {
            args.tokenReady = false;
        } else {
            args.tokenReady = true;
            args.oAuth2Client.setCredentials(JSON.parse(token));
        }
        callback(null, args);
    });
};

var makeToken = function (args, callback) {
    if (!args.tokenReady) {
        if (!args.automated) {
            const authUrl = args.oAuth2Client.generateAuthUrl({
                access_type: 'offline',
                scope: SCOPES,
            });
            console.log('Authorize this app by visiting this url:', authUrl);
            const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

            rl.question('Enter the code from that page here: ', (code) => {
                rl.close();
                args.oAuth2Client.getToken(code, (err, token) => {
                    if (err) {
                        console.error('Error retrieving access token', err);
                        callback(err, args);
                        return;
                    }
                    args.oAuth2Client.setCredentials(token);
                    // Store the token to disk for later program executions
                    fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                        if (err) {
                            console.error(err);
                            callback(err, args);
                            return;
                        }
                        console.log('Token stored to', TOKEN_PATH);
                        callback(null, args);
                        return;
                    });

                });
            });
        } else {
            callback("Token does not exist!", args);
        }
    } else {
        callback(null, args);
    }
};

var clearDate = function (sheet, callback) {
    sheet.service.spreadsheets.values.clear({
        spreadsheetId: sheet.id,
        range: sheet.date_range
    }, (err, res) => {
        callback(null, sheet);
    });
};

var updateDate = function (sheet, callback) {
    console.log(sheet.id, sheet.name, "날짜/요일 입력 요청");
    sheet.service.spreadsheets.values.append({
        spreadsheetId: sheet.id,
        range: sheet.date_range,
        valueInputOption: 'USER_ENTERED',
        resource: { values: sheet.date_value }
    }, (err, res) => {
        if (!err) {
            console.log(sheet.id, sheet.name, "날짜/요일 입력 완료");
        }
        callback(err, sheet);
    });
};

var tracePath = function (sheet, callback) {
    if (sheet.path_value) {
        callback(null, sheet);
        return;
    }

    var option = {
        uri: 'https://api2.sktelecom.com/tmap/routes?version=1',
        method: 'POST',
        form: {
            startX: sheet.start['lon'],
            startY: sheet.start['lat'],
            endX: sheet.end['lon'],
            endY: sheet.end['lat'],
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
            'appKey': sheet.appKey
        },
        jar: true,
        json: true,
        gzip: true,
    };

    console.log(sheet.id, sheet.name, "경로 예상 시간 측정 요청", sheet.path_value);
    request(option, function (err, res, body) {
        if (!err && body && body.features && body.features.length > 0 && body.features[0].properties && body.features[0].properties.totalTime) {
            sheet.path_value = body.features[0].properties.totalTime;
            console.log(sheet.id, sheet.name, "경로 예상 시간 측정 완료", sheet.path_value);
        }

        callback(err, sheet);
    });
};

var foundLocation = function (callback) {
    var option = {
        uri: 'https://apis.openapi.sk.com/tmap/pois?version=1',
        method: 'GET',
        qs: {
            searchKeyword: "네이버랩스",
            reqCoordType: "WGS84GEO",
            resCoordType: "WGS84GEO",
        },
        headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'ko,en-US;q=0.8,en;q=0.6',
            'appKey': config.get('tmap').appKey
        },
        jar: true,
        json: true,
        gzip: true,
    };

    request(option, function (err, res, body) {
        if (!err && body) {
            console.log(JSON.stringify(body, null, "  "));
        }

        callback(err);
    });
};

var updateTime = function (sheet, callback) {
    console.log(sheet.id, sheet.name, "시간 입력 요청");
    if (sheet.path_value > 0) {
        sheet.service.spreadsheets.values.update({
            spreadsheetId: sheet.id,
            range: sheet.path_range,
            valueInputOption: 'USER_ENTERED',
            resource: { values: [[sheet.path_value / (24.0 * 60 * 60)]] }
        }, (err, res) => {
            if (!err) {
                console.log(sheet.id, sheet.name, "시간 입력 완료", res.data.updatedRange, ":", res.statusText);
            }

            callback(err, sheet);
        });
    } else {
        console.log(sheet.id, sheet.name, "시간 측정 실패");
        callback("시간 측정 실패", sheet);
    }
};

var updateViewer = function (args, callback) {
    console.log(viewer_sheets.data, "보기 전용 입력");
    const service = google.sheets({ version: 'v4', auth: args.oAuth2Client });
    service.spreadsheets.values.get({
        spreadsheetId: viewer_sheets.data,
        range: '요약!A:G'
    }, (err, res) => {
        if (!err) {
            console.log(res.data);
            var read_values = res.data.values

            service.spreadsheets.values.clear({
                spreadsheetId: viewer_sheets.viewer,
                range: '요약!A:G',
            }, (err, res) => {
                if (!err) {
                    service.spreadsheets.values.update({
                        spreadsheetId: viewer_sheets.viewer,
                        range: '요약!A:G',
                        valueInputOption: 'RAW',
                        resource: { values: read_values }
                    }, (err, res) => {
                        if (!err) {
                            console.log("보기 전용 입력 완료");
                        }

                        callback(err, args);
                    });
                } else {
                    callback(err, args);
                }
            });
        } else {
            console.log(err);
            callback(err, args);
        }
    });
};

exports.handler = function (event, context, callback) {
    async.waterfall([
        function (callback) {
            const { client_secret, client_id, redirect_uris } = config.get('installed');
            callback(null, {
                oAuth2Client: new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]),
                automated: true,
                tokenReady: false,
                sheets: target_sheets
            });
        },
        authorize,
        makeToken,
        function (args, callback) {
            const service = google.sheets({ version: 'v4', auth: args.oAuth2Client });

            async.each(args.sheets, function (sheet, callback) {
                var start_date = luxon.DateTime.fromISO(sheet.since).setZone('Asia/Seoul');
                var today = luxon.DateTime.local().setZone('Asia/Seoul');
                var row = Math.floor((today - start_date) / 24 / 60 / 60 / 1000) + 2;

                sheet.service = service;
                sheet.date_range = util.format('%s!A%d:B%d', sheet.name, row, row);
                sheet.date_value = [[today.toFormat("yyyy-MM-dd"), weekdays[today.weekday - 1]]];
                sheet.path_range = util.format('%s!%s%d', sheet.name, columns[Math.floor((today.hour * 60 + today.minute) / time_period) + 2], row);
                sheet.path_value = 0;

                async.waterfall([
                    function (callback) {
                        callback(null, sheet);
                    },
                    clearDate,
                    updateDate,
                    tracePath,
                    tracePath,
                    tracePath,
                    updateTime,
                ], function (err) {
                    if (err) {
                        console.log(err);
                    }
                    callback(err);
                });
            }, function (err) {
                if (err) {
                    console.log(err);
                }
                callback(err, args);
            });
        },
        //updateViewer,
    ], function (err, result) {
        if (err) {
            console.log(err);
        }
    });
};
