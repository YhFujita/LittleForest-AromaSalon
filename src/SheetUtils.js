var SheetUtils = (function () {
    // var SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'); // Move inside function to ensure fresh fetch and safety

    var SHEET_NAME_RESERVATIONS = '予約一覧';
    var SHEET_NAME_MENU = 'メニュー設定';
    var SHEET_NAME_SLOTS = '予約可能日時';
    var CACHE_KEY_MENU = 'MENU_CACHE';
    var CACHE_KEY_SLOTS = 'SLOTS_CACHE';

    function getSheet(name) {
        var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
        if (id) id = id.trim();

        var ss;
        if (id) {
            try {
                ss = SpreadsheetApp.openById(id);
            } catch (e) {
                console.error('Error opening Spreadsheet with ID: ' + id, e);
                // Fallback to active spreadsheet if ID is invalid? 
                // However, user specifically set an ID, so it's better to fail or warn. 
                // But for now, let's throw a clearer error.
                throw new Error('スクリプトプロパティで指定されたスプレッドシートを開けませんでした。IDが正しいか確認してください (ID: ' + id + ') エラー: ' + e.message);
            }
        } else {
            // Not set -> use active
            try {
                ss = SpreadsheetApp.getActiveSpreadsheet();
            } catch (e) {
                throw new Error('アクティブなスプレッドシートを取得できませんでした。スタンドアロンスクリプトの場合はSPREADSHEET_IDを設定してください。');
            }
        }

        if (!ss) throw new Error('スプレッドシートが見つかりません。');

        var sheet = ss.getSheetByName(name);

        // Auto-create sheets if missing
        if (!sheet) {
            sheet = ss.insertSheet(name);
            if (name === SHEET_NAME_RESERVATIONS) {
                // Modified header to include '金額'
                sheet.appendRow(['日時', '予約ID', '予約者名', 'メニュー', '金額', '希望日時', '電話番号', '備考', 'ステータス']);
            } else if (name === SHEET_NAME_MENU) {
                sheet.appendRow(['ID', 'メニュー名', '価格', '所要時間(分)', '説明', '表示順']);
                // Add sample data
                sheet.appendRow(['basic', 'ベーシックコース', '6000', '60', '基本のコースです', '1']);
                sheet.appendRow(['premium', 'プレミアムコース', '9000', '90', '充実のコースです', '2']);
            } else if (name === SHEET_NAME_SLOTS) {
                sheet.appendRow(['日時', 'ステータス']);
            }
        }
        return sheet;
    }

    function formatDateJP(date) {
        var days = ['日', '月', '火', '水', '木', '金', '土'];
        var y = date.getFullYear();
        var m = date.getMonth() + 1;
        var d = date.getDate();
        var day = days[date.getDay()];
        var H = ('0' + date.getHours()).slice(-2);
        var M = ('0' + date.getMinutes()).slice(-2);
        return y + '年' + m + '月' + d + '日(' + day + ') ' + H + ':' + M;
    }

    return {
        appendReservation: function (data) {
            var sheet = getSheet(SHEET_NAME_RESERVATIONS);
            var id = Utilities.getUuid();
            var timestamp = new Date();
            var formattedTimestamp = formatDateJP(timestamp);

            // Convert normalized input datetime (YYYY/MM/DD HH:mm) to Display Format for consistent sheet storage
            var bookingDate = new Date(data.datetime);
            var formattedBookingDate = formatDateJP(bookingDate);

            // Lookup Price
            var price = '';
            var menuItems = this.getMenuItems();
            var selectedMenu = menuItems.find(function (item) { return item.id === data.menu; });
            if (selectedMenu) {
                price = selectedMenu.price;
            }

            sheet.appendRow([
                formattedTimestamp, // Recorded at
                id,
                data.name,
                data.menu,
                price, // Saved Price
                formattedBookingDate, // Desired Date (Japanese Format)
                data.phone,
                data.notes,
                '受付'
            ]);
            return id;
        },

        getMenuItems: function () {
            var scriptProperties = PropertiesService.getScriptProperties();
            var cachedMenu = scriptProperties.getProperty(CACHE_KEY_MENU);
            if (cachedMenu) return JSON.parse(cachedMenu);
            return this.updateMenuCache();
        },

        updateMenuCache: function () {
            var sheet = getSheet(SHEET_NAME_MENU);
            var data = sheet.getDataRange().getValues();
            var headers = data.shift();

            var menuItems = data.map(function (row) {
                return {
                    id: row[0],
                    name: row[1],
                    price: row[2],
                    duration: row[3],
                    description: row[4],
                    order: row[5]
                };
            }).filter(function (item) { return item.id && item.name; })
                .sort(function (a, b) { return a.order - b.order; });

            PropertiesService.getScriptProperties().setProperty(CACHE_KEY_MENU, JSON.stringify(menuItems));
            return menuItems;
        },

        getAvailableSlots: function () {
            // Try Cache First
            var scriptProperties = PropertiesService.getScriptProperties();
            var cachedSlots = scriptProperties.getProperty(CACHE_KEY_SLOTS);
            if (cachedSlots) {
                return JSON.parse(cachedSlots);
            }
            return this.updateSlotsCache();
        },

        updateSlotsCache: function () {
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var data = sheet.getDataRange().getValues();
            var headers = data.shift();

            // Filter for "空き" slots
            var slots = data.filter(function (row) {
                return row[1] === '空き';
            }).map(function (row) {
                var date = new Date(row[0]);
                return {
                    value: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
                    display: formatDateJP(date)
                };
            });

            slots.sort(function (a, b) {
                return a.value.localeCompare(b.value);
            });

            PropertiesService.getScriptProperties().setProperty(CACHE_KEY_SLOTS, JSON.stringify(slots));
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
                    // Update cache immediately
                    this.updateSlotsCache();
                    return true; // Success
                }
            }
            return false; // Slot not found or invalid
        },

        addSlots: function (slots) {
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var lastRow = sheet.getLastRow();
            var timeZone = Session.getScriptTimeZone();

            var rowsToAdd = slots.map(function (slot) {
                var date = new Date(slot);
                if (isNaN(date.getTime())) {
                    // Check if valid date
                    return [slot, '空き'];
                }
                // Save as Date object for better data handling
                return [date, '空き'];
            });

            if (rowsToAdd.length > 0) {
                var range = sheet.getRange(lastRow + 1, 1, rowsToAdd.length, 2);
                // Set format to "yyyy年M月d日(ddd) HH:mm" for display
                range.setNumberFormat('yyyy"年"M"月"d"日"(ddd) HH:mm');
                range.setValues(rowsToAdd);
            }

            return this.updateSlotsCache();
        }
    };
})();
