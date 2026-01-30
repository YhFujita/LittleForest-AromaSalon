var SheetUtils = (function () {
    var SPREADSHEET_ID = ''; // ユーザーに設定してもらうか、プロパティスクリプトで管理
    var SHEET_NAME_RESERVATIONS = '予約一覧';
    var SHEET_NAME_MENU = 'メニュー設定';
    var SHEET_NAME_SLOTS = '予約可能日時';
    var CACHE_KEY_MENU = 'MENU_CACHE';

    function getSheet(name) {
        var ss = SPREADSHEET_ID
            ? SpreadsheetApp.openById(SPREADSHEET_ID)
            : SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(name);

        // Auto-create sheets if missing
        if (!sheet) {
            sheet = ss.insertSheet(name);
            if (name === SHEET_NAME_RESERVATIONS) {
                sheet.appendRow(['日時', '予約ID', '予約者名', 'メニュー', '希望日時', '電話番号', '備考', 'ステータス']);
            } else if (name === SHEET_NAME_MENU) {
                sheet.appendRow(['ID', 'メニュー名', '価格', '所要時間(分)', '説明', '表示順']);
                // Add sample data
                sheet.appendRow(['basic', 'ベーシックコース', '6000', '60', '基本のコースです', '1']);
                sheet.appendRow(['premium', 'プレミアムコース', '9000', '90', '充実のコースです', '2']);
            } else if (name === SHEET_NAME_SLOTS) {
                sheet.appendRow(['日時', 'ステータス']);
                // Add sample slots
                var now = new Date();
                now.setDate(now.getDate() + 1); // Tomorrow
                now.setHours(10, 0, 0, 0);
                sheet.appendRow([now, '空き']);
                now.setHours(13, 0, 0, 0);
                sheet.appendRow([now, '空き']);
            }
        }
        return sheet;
    }

    return {
        appendReservation: function (data) {
            var sheet = getSheet(SHEET_NAME_RESERVATIONS);
            var id = Utilities.getUuid();
            var timestamp = new Date();

            sheet.appendRow([
                timestamp,
                id,
                data.name,
                data.menu,
                data.datetime,
                data.phone,
                data.notes,
                '受付' // Initial status
            ]);
            return id;
        },

        getMenuItems: function () {
            // Try reading from PropertyService (Cache) first for speed
            var scriptProperties = PropertiesService.getScriptProperties();
            var cachedMenu = scriptProperties.getProperty(CACHE_KEY_MENU);

            if (cachedMenu) {
                return JSON.parse(cachedMenu);
            }

            // If no cache, read from sheet and update cache
            return this.updateMenuCache();
        },

        updateMenuCache: function () {
            var sheet = getSheet(SHEET_NAME_MENU);
            var data = sheet.getDataRange().getValues();
            var headers = data.shift(); // Remove header row

            // Convert to array of objects
            var menuItems = data.map(function (row) {
                return {
                    id: row[0],
                    name: row[1],
                    price: row[2],
                    duration: row[3],
                    description: row[4],
                    order: row[5]
                };
            }).filter(function (item) {
                // Filter out empty rows
                return item.id && item.name;
            }).sort(function (a, b) {
                return a.order - b.order;
            });

            // Save to Script Properties
            var scriptProperties = PropertiesService.getScriptProperties();
            scriptProperties.setProperty(CACHE_KEY_MENU, JSON.stringify(menuItems));

            return menuItems;
        },

        getAvailableSlots: function () {
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var data = sheet.getDataRange().getValues();
            var headers = data.shift();

            // Filter for "空き" slots
            var slots = data.filter(function (row) {
                return row[1] === '空き';
            }).map(function (row) {
                return Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
            });

            return slots;
        },

        reserveSlot: function (targetDatetimeStr) {
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var data = sheet.getDataRange().getValues();
            var timeZone = Session.getScriptTimeZone();

            // Find row to update
            for (var i = 1; i < data.length; i++) { // Skip header
                var cellDate = new Date(data[i][0]);
                var cellStatus = data[i][1];
                var cellDateStr = Utilities.formatDate(cellDate, timeZone, 'yyyy/MM/dd HH:mm');

                if (cellDateStr === targetDatetimeStr) {
                    if (cellStatus !== '空き') {
                        return false; // Already reserved
                    }
                    // Update status to '予約済'
                    sheet.getRange(i + 1, 2).setValue('予約済');
                    return true; // Success
                }
            }
            return false; // Slot not found or invalid
        }
    };
})();
