var SheetUtils = (function () {
    // var SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'); // Move inside function to ensure fresh fetch and safety

    var SHEET_NAME_RESERVATIONS = '予約一覧';
    var SHEET_NAME_MENU = 'メニュー設定';
    var SHEET_NAME_SLOTS = '予約可能日時';
    var SHEET_NAME_SETTINGS = '設定';
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
                // Modified header to include '金額', swapped Date columns
                sheet.appendRow(['希望日時', '予約ID', '予約者名', 'メニュー', '金額', '送信日時', '電話番号', '備考', 'ステータス', 'GoogleEventID']);
            } else if (name === SHEET_NAME_MENU) {
                sheet.appendRow(['ID', 'メニュー名', '価格', '所要時間(分)', '説明', '表示順', 'カテゴリタイトル']);
                // Add sample data
                sheet.appendRow(['basic', 'ベーシックコース', '6000', '60', '基本のコースです', '1']);
                sheet.appendRow(['premium', 'プレミアムコース', '9000', '90', '充実のコースです', '2']);
            } else if (name === SHEET_NAME_SLOTS) {
                sheet.appendRow(['日時', 'ステータス']);
            } else if (name === SHEET_NAME_SETTINGS) {
                sheet.appendRow(['キー', '値']);
                var defaultTemplate =
                    '{{name}}様、ご予約ありがとうございます。\n' +
                    '以下の内容で承りました。\n\n' +
                    '■日時: {{date}}\n' +
                    '■メニュー: {{menu}}\n' +
                    '■金額: {{price}}円\n\n' +
                    'ご来店をお待ちしております。';
                sheet.appendRow(['LINE_MESSAGE_TEMPLATE', defaultTemplate]);
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
                formattedBookingDate, // Desired Date (A Column)
                id,
                data.name,
                data.menu,
                price, // Saved Price
                formattedTimestamp, // Recorded at (F Column)
                "'" + data.phone,
                data.notes,
                '受付',
                data.googleEventId || '' // GoogleEventID (J Column)
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
                    id: String(row[0]), // Ensure ID is string
                    name: row[1],
                    price: row[2],
                    duration: row[3],
                    description: row[4],
                    duration: row[3],
                    description: row[4],
                    order: row[5],
                    section: row[6] || '' // Category/Section Title
                };
            }).filter(function (item) { return item.id && item.name; })
                .sort(function (a, b) {
                    var orderA = parseInt(a.order) || 9999;
                    var orderB = parseInt(b.order) || 9999;
                    return orderA - orderB;
                });

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

            // 1. Get existing dates to prevent duplicates
            var existingData = sheet.getDataRange().getValues();
            var existingTimes = {}; // Use object for O(1) lookup
            for (var i = 1; i < existingData.length; i++) { // Skip header
                var d = new Date(existingData[i][0]);
                if (!isNaN(d.getTime())) {
                    existingTimes[d.getTime()] = true;
                }
            }

            var rowsToAdd = [];
            slots.forEach(function (slot) {
                var date = new Date(slot);
                // Check validity and duplication
                if (!isNaN(date.getTime()) && !existingTimes[date.getTime()]) {
                    rowsToAdd.push([date, '空き']);
                }
            });

            if (rowsToAdd.length > 0) {
                var lastRow = sheet.getLastRow();
                var range = sheet.getRange(lastRow + 1, 1, rowsToAdd.length, 2);
                range.setNumberFormat('yyyy"年"M"月"d"日"(ddd) HH:mm');
                range.setValues(rowsToAdd);
            }

            // 2. Auto-sort: Status (Desc) -> Date (Asc)
            // '空き' (Open) > '予約済' (Reserved) because '空' (U+7A7A) > '予' (U+4E88)
            var totalRows = sheet.getLastRow();
            if (totalRows > 1) {
                // Sort range excluding header (row 1)
                sheet.getRange(2, 1, totalRows - 1, 2).sort([
                    { column: 2, ascending: false }, // Status: Open first
                    { column: 1, ascending: true }   // Date: Chronological
                ]);
            }

            return this.updateSlotsCache();
        },

        // --- Admin Functions ---

        deleteSlot: function (targetDatetimeStr) {
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var data = sheet.getDataRange().getValues();
            var timeZone = Session.getScriptTimeZone();

            for (var i = 1; i < data.length; i++) {
                var cellDate = new Date(data[i][0]);
                var cellDateStr = Utilities.formatDate(cellDate, timeZone, 'yyyy/MM/dd HH:mm');

                // Compare with target datetime (assuming target passed as normalized string)
                if (cellDateStr === targetDatetimeStr) {
                    sheet.deleteRow(i + 1);
                    this.updateSlotsCache();
                    return true;
                }
            }
            return false;
        },

        updateSlotStatus: function (targetDatetimeStr, newStatus) {
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var data = sheet.getDataRange().getValues();
            var timeZone = Session.getScriptTimeZone();

            for (var i = 1; i < data.length; i++) {
                var cellDate = new Date(data[i][0]);
                var cellDateStr = Utilities.formatDate(cellDate, timeZone, 'yyyy/MM/dd HH:mm');

                if (cellDateStr === targetDatetimeStr) {
                    sheet.getRange(i + 1, 2).setValue(newStatus);
                    this.updateSlotsCache();
                    return true;
                }
            }
            return false;
        },

        updateSlotDatetime: function (currentDatetimeStr, newDateObj) {
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var data = sheet.getDataRange().getValues();
            var timeZone = Session.getScriptTimeZone();
            var newDateStr = Utilities.formatDate(newDateObj, timeZone, 'yyyy/MM/dd HH:mm');

            // 1. Check for duplicates (exclude the current slot being edited if it was unchanged, but here we check consistency)
            for (var i = 1; i < data.length; i++) {
                var d = new Date(data[i][0]);
                if (!isNaN(d.getTime())) {
                    var dStr = Utilities.formatDate(d, timeZone, 'yyyy/MM/dd HH:mm');
                    if (dStr === newDateStr) {
                        return { success: false, message: 'その日時は既に存在します' };
                    }
                }
            }

            // 2. Find and update
            for (var i = 1; i < data.length; i++) {
                var cellDate = new Date(data[i][0]);
                var cellDateStr = Utilities.formatDate(cellDate, timeZone, 'yyyy/MM/dd HH:mm');

                if (cellDateStr === currentDatetimeStr) {
                    sheet.getRange(i + 1, 1).setValue(newDateObj);

                    // 3. Sort
                    var totalRows = sheet.getLastRow();
                    if (totalRows > 1) {
                        sheet.getRange(2, 1, totalRows - 1, 2).sort([
                            { column: 2, ascending: false },
                            { column: 1, ascending: true }
                        ]);
                    }

                    this.updateSlotsCache();
                    return { success: true };
                }
            }
            return { success: false, message: '対象の予約枠が見つかりませんでした' };
        },

        saveMenuItem: function (item) {
            var sheet = getSheet(SHEET_NAME_MENU);
            var data = sheet.getDataRange().getValues();
            var updated = false;

            // Try to update existing
            if (item.id) {
                for (var i = 1; i < data.length; i++) {
                    if (data[i][0] == item.id) {
                        // Update row (now 7 cols)
                        var range = sheet.getRange(i + 1, 1, 1, 7);
                        range.setValues([[
                            item.id,
                            item.name,
                            item.price,
                            item.duration,
                            item.description,
                            item.order,
                            item.section
                        ]]);
                        // Format Price (Column 3 is relative index 2? No, getRange(row, 1, 1, 6) -> 3rd cell is col index 3 in sheet)
                        // sheet.getRange(row, col) -> Price is 3rd column
                        sheet.getRange(i + 1, 3).setNumberFormat('#,##0');
                        updated = true;
                        break;
                    }
                }
            }

            // Insert new if not updated
            if (!updated) {
                item.id = item.id || Utilities.getUuid();
                sheet.appendRow([
                    item.id,
                    item.name,
                    item.price,
                    item.duration,
                    item.description,
                    item.order,
                    item.section
                ]);
                // Format the newly added row's price column
                var lastRow = sheet.getLastRow();
                sheet.getRange(lastRow, 3).setNumberFormat('#,##0');
            }

            return this.updateMenuCache();
        },

        reorderMenuItems: function (ids) {
            var sheet = getSheet(SHEET_NAME_MENU);
            var data = sheet.getDataRange().getValues();
            var idMap = {}; // Map ID to Row Index for O(1)

            // Build map (skip header)
            for (var i = 1; i < data.length; i++) {
                idMap[String(data[i][0])] = i + 1; // 1-based index
            }

            // Update order for each ID in the new list
            ids.forEach(function (id, index) {
                var rowIndex = idMap[id];
                if (rowIndex) {
                    // Update 'order' column (Column 6) with new index (1-based or 0-based doesn't matter as long as it sorts)
                    // We use (index + 1) * 10 to leave room if needed
                    sheet.getRange(rowIndex, 6).setValue((index + 1) * 10);
                }
            });

            return this.updateMenuCache();
        },

        getReservations: function () {
            var sheet = getSheet(SHEET_NAME_RESERVATIONS);
            var data = sheet.getDataRange().getValues();
            var timeZone = Session.getScriptTimeZone();
            var results = [];

            // Skip header (row 1)
            for (var i = 1; i < data.length; i++) {
                var row = data[i];
                // Check if date is valid
                var d = new Date(row[0]);
                if (isNaN(d.getTime())) {
                    // Try parsing Japanese format: "yyyy年M月d日(day) H:m"
                    // Regex to capture digits. Ignore the day of week char.
                    var match = String(row[0]).match(/(\d+)年(\d+)月(\d+)日\s*\(.\)\s*(\d+):(\d+)/);
                    if (match) {
                        d = new Date(
                            parseInt(match[1], 10),
                            parseInt(match[2], 10) - 1,
                            parseInt(match[3], 10),
                            parseInt(match[4], 10),
                            parseInt(match[5], 10)
                        );
                    }
                }
                if (isNaN(d.getTime())) continue;

                // Columns: 0:Date, 1:ID, 2:Name, 3:MenuID, 4:Price, 5:Timestamp, 6:Phone, 7:Notes, 8:Status, 9:GoogleEventID
                results.push({
                    date: Utilities.formatDate(d, timeZone, 'yyyy/MM/dd HH:mm'),
                    displayDate: row[0], // Use original string as display or re-format? Original is fine if consistent.
                    id: row[1],
                    name: row[2],
                    menuId: row[3],
                    price: row[4],
                    phone: row[6],
                    notes: row[7],
                    status: row[8],
                    googleEventId: row[9] || ''
                });
            }

            // Sort by Date Descending (Newest first) or Ascending (Future first)?
            // Usually Admin wants to see upcoming. Ascending.
            results.sort(function (a, b) {
                return a.date.localeCompare(b.date);
            });

            return results;
        },

        deleteMenuItem: function (id) {
            var sheet = getSheet(SHEET_NAME_MENU);
            var data = sheet.getDataRange().getValues();

            for (var i = 1; i < data.length; i++) {
                if (data[i][0] == id) {
                    sheet.deleteRow(i + 1);
                    this.updateMenuCache();
                    return true;
                }
            }
            return false;
        },

        setReservationGoogleEventId: function (reservationId, eventId) {
            var sheet = getSheet(SHEET_NAME_RESERVATIONS);
            var data = sheet.getDataRange().getValues();

            for (var i = 1; i < data.length; i++) {
                if (data[i][1] == reservationId) { // Column B is ID
                    sheet.getRange(i + 1, 10).setValue(eventId); // Column J is GoogleEventID
                    return true;
                }
            }
            return false;
        },

        getReservation: function (reservationId) {
            var sheet = getSheet(SHEET_NAME_RESERVATIONS);
            var data = sheet.getDataRange().getValues();

            for (var i = 1; i < data.length; i++) {
                if (data[i][1] == reservationId) {
                    var d = new Date(data[i][0]);
                    if (isNaN(d.getTime())) {
                        var match = String(data[i][0]).match(/(\d+)年(\d+)月(\d+)日\s*\(.\)\s*(\d+):(\d+)/);
                        if (match) {
                            d = new Date(
                                parseInt(match[1], 10),
                                parseInt(match[2], 10) - 1,
                                parseInt(match[3], 10),
                                parseInt(match[4], 10),
                                parseInt(match[5], 10)
                            );
                        }
                    }
                    return {
                        row: i + 1,
                        date: d,
                        id: data[i][1],
                        menu: data[i][3],
                        status: data[i][8],
                        googleEventId: data[i][9]
                    };
                }
            }
            return null;
        },

        cancelReservation: function (reservationId) {
            var sheet = getSheet(SHEET_NAME_RESERVATIONS);
            var data = sheet.getDataRange().getValues();
            var timeZone = Session.getScriptTimeZone();

            for (var i = 1; i < data.length; i++) {
                if (data[i][1] == reservationId) {
                    // 1. Update Status
                    sheet.getRange(i + 1, 9).setValue('キャンセル');

                    // 2. Release Slot
                    var date = new Date(data[i][0]);
                    if (isNaN(date.getTime())) {
                        var match = String(data[i][0]).match(/(\d+)年(\d+)月(\d+)日\s*\(.\)\s*(\d+):(\d+)/);
                        if (match) {
                            date = new Date(
                                parseInt(match[1], 10),
                                parseInt(match[2], 10) - 1,
                                parseInt(match[3], 10),
                                parseInt(match[4], 10),
                                parseInt(match[5], 10)
                            );
                        }
                    }
                    var dateStr = Utilities.formatDate(date, timeZone, 'yyyy/MM/dd HH:mm');
                    this.updateSlotStatus(dateStr, '空き');

                    return {
                        success: true,
                        googleEventId: data[i][9],
                        date: dateStr,
                        menu: data[i][3]
                    };
                }
            }
            return { success: false, message: 'Reservation not found' };
        },

        updateReservation: function (reservationId, newDatetime, newMenuId) {
            var res = this.getReservation(reservationId);
            if (!res) return { success: false, message: 'Reservation not found' };

            var sheet = getSheet(SHEET_NAME_RESERVATIONS);
            var timeZone = Session.getScriptTimeZone();
            var currentDatetimeStr = Utilities.formatDate(res.date, timeZone, 'yyyy/MM/dd HH:mm');
            var newDatetimeStr = Utilities.formatDate(new Date(newDatetime), timeZone, 'yyyy/MM/dd HH:mm');

            // 1. Check if new slot is available (unless it's the same time)
            if (currentDatetimeStr !== newDatetimeStr) {
                if (!this.reserveSlot(newDatetimeStr)) {
                    return { success: false, message: '変更先の日時は既に埋まっています。' };
                }
                // Release old slot
                this.updateSlotStatus(currentDatetimeStr, '空き');
            }

            // 2. Update Reservation Data
            var row = res.row;
            // Update Date (Col A)
            sheet.getRange(row, 1).setValue(formatDateJP(new Date(newDatetime)));

            // Update Menu & Price if changed
            if (newMenuId && newMenuId !== res.menu) {
                var menuItems = this.getMenuItems();
                var selectedMenu = menuItems.find(function (item) { return item.id === newMenuId; });
                if (selectedMenu) {
                    sheet.getRange(row, 4).setValue(newMenuId); // ID stored in col D?
                    // Wait, original appendReservation stores `data.menu` which is ID? 
                    // Looking at appendRow: `data.menu`
                    // Line 101: data.menu. 
                    // So yes, it stores the ID.
                    sheet.getRange(row, 5).setValue(selectedMenu.price);
                }
            }

            return {
                success: true,
                googleEventId: res.googleEventId,
                oldDate: res.date,
                newDate: new Date(newDatetime)
            };
        }
    };

})();
