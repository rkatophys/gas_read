function slackReadText() {
    var token = PropertiesService.getScriptProperties().getProperty("SLACK_TOKEN");
    var driveFolderId = PropertiesService.getScriptProperties().getProperty("DRIVE_ID");

    var options = {
        method: 'GET',
        headers: { 'Authorization': 'Bearer ' + token },
    };
    channelApiUrl = "https://slack.com/api/conversations.list"
    var channelResponse = UrlFetchApp.fetch(channelApiUrl, options);
    var channelData = JSON.parse(channelResponse.getContentText());
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()

    // パブリックなチャンネルそれぞれのシートを作る。
    for (var i = 0; i < channelData.channels.length; i++) {
        var eachChannelData = channelData.channels[i];
        var hasSheetName = false;
        for (const sheet of sheets) {
            var sheetName = sheet.getSheetName()
            if (sheetName == eachChannelData.name) {
                hasSheetName = true;
            }
        }
        if (!hasSheetName) {
            var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
            newSheet.setName(eachChannelData.name);
        }
    }

    // すべてのユーザ情報の取得
    var allUserApiUrl = "https://slack.com/api/users.list";
    var allUserResponse = UrlFetchApp.fetch(allUserApiUrl, options);
    var allUserData = JSON.parse(allUserResponse);

    var allUserDataDict = {}
    for (var i = 0; i < allUserData.members.length; i++) {
        var eachMember = allUserData.members[i];
        allUserDataDict[eachMember.id] = eachMember.profile.real_name;
    }

    for (var i = 0; i < channelData.channels.length; i++) {
        var eachChannelData = channelData.channels[i];
        var eachChannel = eachChannelData.id
        var eachSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(eachChannelData.name)
        assignSheet(token, driveFolderId, eachChannel, eachSheet, allUserDataDict)
    }
}

function assignSheet(token, driveFolderId, channel, sheet, allUserDataDict) {
    sheet.setColumnWidth(1, 30);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 500);
    sheet.setColumnWidth(5, 140);
    sheet.getRange("D:D").setWrap(true)

    // ヘッダーの設定
    if (sheet.getRange("B4").getValue() === "") {
        sheet.getRange("B4").setValue("投稿時刻");
        sheet.getRange("C4").setValue("投稿者");
        sheet.getRange("D4").setValue("投稿内容");
        sheet.getRange("E4").setValue("スレッド開始時刻");
        sheet.getRange("F4").setValue("ファイル名");
        sheet.getRange("G4").setValue("ファイルURL");
    }

    sheet.getRange("B2").setValue("最終更新時刻");
    var lastExecutionTime = sheet.getRange("B3").getValue();

    sheet.getRange("C2").setValue("全ファイルのフォルダ");
    var gurl = sheet.getRange("C3").setValue("https://drive.google.com/drive/u/0/folders/" + driveFolderId);
    gurl.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)


    var apiUrl = "https://slack.com/api/conversations.history?channel=" + channel + "&limit=10000";
    var options = {
        method: 'GET',
        headers: { 'Authorization': 'Bearer ' + token },
    };
    var response = UrlFetchApp.fetch(apiUrl, options);
    var data = JSON.parse(response.getContentText());
    var message = data.messages;

    // 返信も含めたメッセージを作りソート
    for (var i = 0; i < message.length; i++) {
        if (message[i].reply_count > 0) {
            var repliesApiUrl = "https://slack.com/api/conversations.replies?channel=" + channel + "&ts=" + message[i].ts;
            var repliesResponse = UrlFetchApp.fetch(repliesApiUrl, options);
            var repliesData = JSON.parse(repliesResponse.getContentText());
            var repliesMessage = repliesData.messages;

            for (var j = 1; j < repliesData.messages.length; j++) {
                message = message.concat([repliesMessage[j]]);
            }

        }
    }
    message.sort((a, b) => a.ts - b.ts);



    for (var i = 0; i < message.length; i++) {
        // メインメッセージの処理
        processMessage(message[i], token, driveFolderId, sheet, lastExecutionTime, allUserDataDict);
    }

    datecell = sheet.getRange("B3").setValue(new Date());
    datecell.setNumberFormat("yyyy/MM/dd H:mm:ss");
}

function processMessage(message, token, driveFolderId, sheet, lastExecutionTime, allUserDataDict) {
    var options = {
        method: 'GET',
        headers: { 'Authorization': 'Bearer ' + token },
    };
    var messageTimestamp = new Date(message.ts * 1000);
    if (typeof message.thread_ts === "undefined") {
        var messageThreadTimestamp = false;
    } else {
        var messageThreadTimestamp = new Date(message.thread_ts * 1000);
    }

    if (messageTimestamp > lastExecutionTime) {
        // ユーザ情報の取得
        var userApiUrl = "https://slack.com/api/users.info?user=" + message.user;
        var userResponse = UrlFetchApp.fetch(userApiUrl, options);
        if (userResponse.status == 429) {
            let retryAfter = userResponse.headers['retry-after'];
            queue.pause()
            wait(retryAfter * 1000)
            queue.resume()
        }
        var userData = JSON.parse(userResponse.getContentText());
        var nextRow = sheet.getLastRow() + 1;
        var userName = userData.user ? userData.user.real_name : "Unknown User";
        timecell = sheet.getRange(nextRow, 2).setValue(messageTimestamp);
        timecell.setNumberFormat("yyyy/MM/dd H:mm:ss");
        sheet.getRange(nextRow, 3).setValue(userName);

        // メッセージに含まれるIDを名前に変換
        var repText = message.text;
        for (const [key, value] of Object.entries(allUserDataDict)) {
            repText = repText.replaceAll(key, value);
        }

        sheet.getRange(nextRow, 4).setValue(repText);
        tsCell = sheet.getRange(nextRow, 5).setValue(messageThreadTimestamp);
        tsCell.setNumberFormat("yyyy/MM/dd H:mm:ss");

        if (message.files) {
            for (var j = 0; j < message.files.length; j++) {
                var file = message.files[j];

                var fileUrl = file.url_private_download;

                Logger.log("File URL: " + fileUrl);

                if (!fileUrl) {
                    Logger.log("URL not found for file: " + file.name);
                    continue;
                }

                try {
                    var fileResponse = UrlFetchApp.fetch(fileUrl, options);
                    var blob = fileResponse.getBlob().setName(file.name);
                    // 同一ファイル名がある場合は上書きしない。
                    const folder = DriveApp.getFolderById(driveFolderId)
                    const files = folder.getFiles()
                    const fileName = (file.name)
                    let hasFileName = false
                    while (files.hasNext()) {
                        const file = files.next()
                        if (file.getName() === fileName) {
                            hasFileName = true;
                            var savedFile = file;
                            console.log(`${fileName}ファイルが存在します。`);
                            break
                        }
                    }
                    if (!hasFileName) {
                        var savedFile = folder.createFile(blob);
                        console.log(`${fileName}ファイルがないので作成しました。`);
                    }
                    var driveUrl = "https://drive.google.com/file/d/" + savedFile.getId() + "/view";
                    sheet.getRange(4, 6 + 2 * j).setValue("ファイル名" + String(j + 1));
                    sheet.getRange(4, 7 + 2 * j).setValue("ファイルURL" + String(j + 1))
                    sheet.getRange(nextRow, 6 + 2 * j).setValue(file.name);
                    var url = sheet.getRange(nextRow, 7 + 2 * j);
                    url.setValue(driveUrl);
                    url.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

                } catch (error) {
                    Logger.log("Error downloading or saving the file: " + file.name);
                    Logger.log(error.toString());
                }
            }
        }
    }
}

