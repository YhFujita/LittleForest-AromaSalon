var SheetUtils = (function () {
    // Updated: Fix Date Parsing & Add Columns
    // var SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'); 

    var SHEET_NAME_RESERVATIONS = '予約一覧';
    var SHEET_NAME_MENU = 'メニュー設定';
    var SHEET_NAME_SLOTS = '予約可能日時';
    var SHEET_NAME_SETTINGS = '設定';
    var SHEET_NAME_BASIC_SETTINGS = '基本設定';
    var CACHE_KEY_MENU = 'MENU_CACHE';
    var CACHE_KEY_SLOTS = 'SLOTS_CACHE';
    var CACHE_KEY_BASIC_SETTINGS = 'BASIC_SETTINGS_CACHE';

    // Helper to extract HH:mm from Date or String
    function fmtTime(val) {
        if (!val) return '';
        if (val instanceof Date) {
            return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
        }
        return String(val);
    }

    function getSheet(name) {
        var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
        if (id) id = id.trim();

        var ss;
        if (id) {
            try {
                ss = SpreadsheetApp.openById(id);
            } catch (e) {
                console.error('Error opening Spreadsheet with ID: ' + id, e);
                throw new Error('スクリプトプロパティで指定されたスプレッドシートを開けませんでした。IDが正しいか確認してください (ID: ' + id + ') エラー: ' + e.message);
            }
        } else {
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
                // 新しいヘッダー構成
                sheet.appendRow(['希望日時', '予約ID', '予約者名', 'メニュー(ID)', 'メニュー名', '税抜金額', '消費税', '金額(税込)', '送信日時', '電話番号', '備考', 'ステータス', 'GoogleEventID']);
            } else if (name === SHEET_NAME_MENU) {
                sheet.appendRow(['ID', 'メニュー名', '価格', '所要時間(分)', '説明', '表示順', 'カテゴリタイトル']);
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
                sheet.appendRow(['SLOT_INTERVAL', '60']); // Default interval
            } else if (name === SHEET_NAME_BASIC_SETTINGS) {
                // 新しい構造: [曜日ID, 曜日名, 営業ステータス, シフト設定]
                // シフト設定: "10:00-13:00,14:00-18:00" のようなカンマ区切り文字列
                sheet.appendRow(['曜日ID', '曜日名', '営業ステータス', 'シフト設定']);
                var days = ['日', '月', '火', '水', '木', '金', '土'];
                for (var i = 0; i < 7; i++) {
                    sheet.appendRow([i, days[i], '営業', '10:00-20:00']);
                }
            }
        } else if (name === SHEET_NAME_RESERVATIONS) {
            // ... (Existing reservation migration logic) ...
            var header = sheet.getRange(1, 1, 1, 10).getValues()[0];
            if (header[4] === '金額') {
                sheet.insertColumnsAfter(4, 3);
                sheet.getRange(1, 4).setValue('メニュー(ID)');
                sheet.getRange(1, 5).setValue('メニュー名');
                sheet.getRange(1, 6).setValue('税抜金額');
                sheet.getRange(1, 7).setValue('消費税');
                sheet.getRange(1, 8).setValue('金額(税込)');
            }
        } else if (name === SHEET_NAME_BASIC_SETTINGS) {
            // Migration: Convert old 7-column format to new 4-column format
            // Old: [ID, Name, Status, Start, End, BreakStart, BreakEnd]
            // New: [ID, Name, Status, Shifts]
            var lastCol = sheet.getLastColumn();
            if (lastCol >= 7) {
                var data = sheet.getDataRange().getValues();
                var header = data[0];
                // Check if index 3 is '開始時間' (Old)
                if (header[3] === '開始時間') {
                    var newData = [];
                    newData.push(['曜日ID', '曜日名', '営業ステータス', 'シフト設定']);

                    for (var i = 1; i < data.length; i++) {
                        var row = data[i];
                        var status = row[2];
                        if (status !== '営業' && status !== '休業') status = '営業';

                        var shifts = [];
                        if (status === '営業') {
                            var start = row[3] ? fmtTime(row[3]) : '10:00';
                            var end = row[4] ? fmtTime(row[4]) : '20:00';
                            var bStart = row[5] ? fmtTime(row[5]) : '';
                            var bEnd = row[6] ? fmtTime(row[6]) : '';

                            if (bStart && bEnd) {
                                shifts.push(start + '-' + bStart);
                                shifts.push(bEnd + '-' + end);
                            } else {
                                shifts.push(start + '-' + end);
                            }
                        }
                        newData.push([row[0], row[1], status, shifts.join(',')]);
                    }

                    sheet.clear();
                    sheet.getRange(1, 1, newData.length, 4).setValues(newData);
                }
            }
        }
        return sheet;
    }

    // ... (Existing formatDateJP, parseDatetimeString) ...
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

    function parseDatetimeString(datetimeStr) {
        if (!datetimeStr) return null;
        var normalized = String(datetimeStr).replace(/\//g, '-');
        var match = normalized.match(/(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{2})/);
        if (match) {
            return new Date(
                parseInt(match[1], 10),
                parseInt(match[2], 10) - 1,
                parseInt(match[3], 10),
                parseInt(match[4], 10),
                parseInt(match[5], 10)
            );
        }
        var d = new Date(datetimeStr);
        if (!isNaN(d.getTime())) return d;
        return null;
    }

    return {
        // ... (Existing functions: appendReservation, getMenuItems, updateMenuCache, getAvailableSlots, updateSlotsCache, reserveSlot, addSlots) ...

        getBasicSettings: function () {
            var scriptProperties = PropertiesService.getScriptProperties();
            var cached = scriptProperties.getProperty(CACHE_KEY_BASIC_SETTINGS);
            if (cached) return JSON.parse(cached);
            return this.updateBasicSettingsCache();
        },

        updateBasicSettingsCache: function () {
            var sheet = getSheet(SHEET_NAME_BASIC_SETTINGS);
            var data = sheet.getDataRange().getValues();
            data.shift(); // Remove header

            var settings = data.map(function (row) {
                // row: [0:ID, 1:Name, 2:Status, 3:Shifts]
                var shiftStr = String(row[3] || '');
                var shifts = [];
                if (shiftStr) {
                    var parts = shiftStr.split(',');
                    parts.forEach(function (p) {
                        var times = p.split('-');
                        if (times.length === 2) {
                            shifts.push({
                                start: times[0].trim(),
                                end: times[1].trim()
                            });
                        }
                    });
                }

                return {
                    dayId: row[0],
                    dayName: row[1],
                    status: row[2], // '営業' or '休業'
                    shifts: shifts
                };
            });

            // Get Interval from SETTINGS sheet
            var settingsSheet = getSheet(SHEET_NAME_SETTINGS);
            var settingsData = settingsSheet.getDataRange().getValues();
            var interval = 60; // Default
            for (var i = 1; i < settingsData.length; i++) {
                if (settingsData[i][0] === 'SLOT_INTERVAL') {
                    interval = parseInt(settingsData[i][1], 10) || 60;
                    break;
                }
            }

            var result = {
                weekly: settings,
                interval: interval
            };

            PropertiesService.getScriptProperties().setProperty(CACHE_KEY_BASIC_SETTINGS, JSON.stringify(result));
            return result;
        },

        saveBasicSettings: function (settings) {
            // settings: { weekly: [{dayId, dayName, status, shifts:[{start,end},...]}], interval: 60 }
            var sheet = getSheet(SHEET_NAME_BASIC_SETTINGS);

            // Prepare data
            var weeklyData = settings.weekly.map(function (item) {
                var shiftStrs = [];
                if (item.shifts && Array.isArray(item.shifts)) {
                    item.shifts.forEach(function (s) {
                        if (s.start && s.end) {
                            shiftStrs.push(s.start + '-' + s.end);
                        }
                    });
                }
                return [
                    item.dayId,
                    item.dayName,
                    item.status,
                    shiftStrs.join(',')
                ];
            });

            // Overwrite from row 2
            if (weeklyData.length > 0) {
                // Clear previous content to ensure no leftover columns if any
                sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
                sheet.getRange(2, 1, weeklyData.length, 4).setValues(weeklyData);
            }

            // Update Interval
            var settingsSheet = getSheet(SHEET_NAME_SETTINGS);
            var sData = settingsSheet.getDataRange().getValues();
            var found = false;
            for (var i = 1; i < sData.length; i++) {
                if (sData[i][0] === 'SLOT_INTERVAL') {
                    settingsSheet.getRange(i + 1, 2).setValue(settings.interval);
                    found = true;
                    break;
                }
            }
            if (!found) {
                settingsSheet.appendRow(['SLOT_INTERVAL', settings.interval]);
            }

            return this.updateBasicSettingsCache();
        },

        updateDaySlots: function (targetDateStr, activeSlots) {
            // targetDateStr: "YYYY-MM-DD"
            // activeSlots: ["YYYY-MM-DD HH:mm", ...]
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var data = sheet.getDataRange().getValues();
            var timeZone = Session.getScriptTimeZone();

            // 1. Identify rows to delete (Open slots on that day)
            // We iterate backwards to delete safely
            for (var i = data.length - 1; i >= 1; i--) {
                var rowDate = new Date(data[i][0]);
                if (isNaN(rowDate.getTime())) continue;
                var rowDateStr = Utilities.formatDate(rowDate, timeZone, 'yyyy-MM-dd');

                if (rowDateStr === targetDateStr) {
                    var status = data[i][1];
                    // Only delete '空き'. Keep '予約済' or others.
                    if (status === '空き') {
                        sheet.deleteRow(i + 1);
                    }
                }
            }

            // 2. Add new slots
            if (activeSlots && activeSlots.length > 0) {
                // Need to filter out slots that already exist (e.g. '予約済' ones we kept)
                // Let's re-fetch data to be sure
                var currentData = sheet.getDataRange().getValues();
                var existingTimeSet = {};
                for (var i = 1; i < currentData.length; i++) {
                    var d = new Date(currentData[i][0]);
                    if (!isNaN(d.getTime())) {
                        var tStr = Utilities.formatDate(d, timeZone, 'yyyy/MM/dd HH:mm');
                        existingTimeSet[tStr] = true;
                    }
                }

                var rowsToAdd = [];
                activeSlots.forEach(function (slotStr) {
                    // slotStr should be formatted as needed, assuming input is parseable
                    var sDate = new Date(slotStr);
                    var sDateStr = Utilities.formatDate(sDate, timeZone, 'yyyy/MM/dd HH:mm');

                    if (!existingTimeSet[sDateStr]) {
                        rowsToAdd.push([sDate, '空き']);
                    }
                });

                if (rowsToAdd.length > 0) {
                    var lastRow = sheet.getLastRow();
                    var range = sheet.getRange(lastRow + 1, 1, rowsToAdd.length, 2);
                    range.setNumberFormat('yyyy"年"M"月"d"日"(ddd) HH:mm');
                    range.setValues(rowsToAdd);
                }
            }

            // 3. Sort
            var totalRows = sheet.getLastRow();
            if (totalRows > 1) {
                sheet.getRange(2, 1, totalRows - 1, 2).sort([
                    { column: 2, ascending: false },
                    { column: 1, ascending: true }
                ]);
            }
            return this.updateSlotsCache();
        },

        appendReservation: function (data) {
            var sheet = getSheet(SHEET_NAME_RESERVATIONS);
            var id = Utilities.getUuid();
            var timestamp = new Date();
            var formattedTimestamp = formatDateJP(timestamp);
            var bookingDate = new Date(data.datetime);
            var formattedBookingDate = formatDateJP(bookingDate);

            // Lookup Price & Menu Name
            var price = 0;
            var menuName = '';
            var menuItems = this.getMenuItems();
            var selectedMenu = menuItems.find(function (item) { return item.id === data.menu; });
            if (selectedMenu) {
                price = parseInt(selectedMenu.price, 10) || 0;
                menuName = selectedMenu.name;
            }

            // Calculate Tax
            var taxRate = 0.10;
            var priceExcl = Math.ceil(price / (1 + taxRate)); // 税込から逆算なので切り上げor四捨五入？一般的には 本体 + 消費税 = 税込
            // 税抜 = 税込 / 1.1 -> 端数処理。
            // 例: 1100円 -> 1000円
            // 例: 100円 -> 90.9... -> 91円?
            // 日本の商慣習では「切り捨て」が多いが、税込価格設定の場合は「内税」計算。
            // 国税庁: 総額表示。消費税額 = 支払総額 × 10 / 110 (円未満端数処理は事業者の判断)
            // ここでは「切り捨て」で計算し、残りを本体とするのが安全か、あるいは 本体=round(税込/1.1) か。
            // ここではシンプルに: 税抜 = Math.floor(price / 1.1), 税 = price - 税抜
            var priceExcl = Math.floor(price / 1.1);
            var tax = price - priceExcl;

            sheet.appendRow([
                formattedBookingDate, // A
                id,                   // B
                data.name,            // C
                data.menu,            // D (Menu ID)
                menuName,             // E (Menu Name)
                priceExcl,            // F (Excl Tax)
                tax,                  // G (Tax)
                price,                // H (Incl Tax)
                formattedTimestamp,   // I
                "'" + data.phone,     // J
                data.notes,           // K
                '受付',               // L
                data.googleEventId || '' // M
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
                    id: String(row[0]),
                    name: row[1],
                    price: row[2],
                    duration: row[3],
                    description: row[4],
                    order: row[5],
                    section: row[6] || ''
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

            for (var i = 1; i < data.length; i++) {
                var cellDate = new Date(data[i][0]);
                var cellStatus = data[i][1];
                var cellDateStr = Utilities.formatDate(cellDate, timeZone, 'yyyy/MM/dd HH:mm');

                if (cellDateStr === targetDatetimeStr) {
                    if (cellStatus !== '空き') {
                        return false;
                    }
                    sheet.getRange(i + 1, 2).setValue('予約済');
                    this.updateSlotsCache();
                    return true;
                }
            }
            return false;
        },

        addSlots: function (slots) {
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var existingData = sheet.getDataRange().getValues();
            var existingTimes = {};
            for (var i = 1; i < existingData.length; i++) {
                var d = new Date(existingData[i][0]);
                if (!isNaN(d.getTime())) {
                    existingTimes[d.getTime()] = true;
                }
            }

            var rowsToAdd = [];
            slots.forEach(function (slot) {
                var date = new Date(slot);
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

            var totalRows = sheet.getLastRow();
            if (totalRows > 1) {
                sheet.getRange(2, 1, totalRows - 1, 2).sort([
                    { column: 2, ascending: false },
                    { column: 1, ascending: true }
                ]);
            }
            return this.updateSlotsCache();
        },

        deleteSlot: function (targetDatetimeStr) {
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var data = sheet.getDataRange().getValues();
            var timeZone = Session.getScriptTimeZone();

            for (var i = 1; i < data.length; i++) {
                var cellDate = new Date(data[i][0]);
                var cellDateStr = Utilities.formatDate(cellDate, timeZone, 'yyyy/MM/dd HH:mm');

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

            for (var i = 1; i < data.length; i++) {
                var d = new Date(data[i][0]);
                if (!isNaN(d.getTime())) {
                    var dStr = Utilities.formatDate(d, timeZone, 'yyyy/MM/dd HH:mm');
                    if (dStr === newDateStr) {
                        return { success: false, message: 'その日時は既に存在します' };
                    }
                }
            }

            for (var i = 1; i < data.length; i++) {
                var cellDate = new Date(data[i][0]);
                var cellDateStr = Utilities.formatDate(cellDate, timeZone, 'yyyy/MM/dd HH:mm');

                if (cellDateStr === currentDatetimeStr) {
                    sheet.getRange(i + 1, 1).setValue(newDateObj);
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

            if (item.id) {
                for (var i = 1; i < data.length; i++) {
                    if (data[i][0] == item.id) {
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
                        sheet.getRange(i + 1, 3).setNumberFormat('#,##0');
                        updated = true;
                        break;
                    }
                }
            }

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
                var lastRow = sheet.getLastRow();
                sheet.getRange(lastRow, 3).setNumberFormat('#,##0');
            }
            return this.updateMenuCache();
        },

        reorderMenuItems: function (ids) {
            var sheet = getSheet(SHEET_NAME_MENU);
            var data = sheet.getDataRange().getValues();
            var idMap = {};
            for (var i = 1; i < data.length; i++) {
                idMap[String(data[i][0])] = i + 1;
            }
            ids.forEach(function (id, index) {
                var rowIndex = idMap[id];
                if (rowIndex) {
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

            for (var i = 1; i < data.length; i++) {
                var row = data[i];
                var d = new Date(row[0]);
                if (isNaN(d.getTime())) {
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

                // Columns: 
                // 0:Date, 1:ID, 2:Name, 3:MenuID, 4:MenuName, 5:Excl, 6:Tax, 7:Incl, 8:Time, 9:Phone, 10:Notes, 11:Status, 12:EventID
                // 注意: 旧データで列が不足している場合のガードが必要かもしれないが、
                // getSheetで列挿入しているので data.length は増えているはず?
                // data = getValues() はシート全体のデータを取るので、行によって列数が違うことはない(空白になる)

                // 互換性: 旧データのこの列は空白かもしれない
                var price = row[7]; // New Price Column (Incl)
                if (price === undefined || price === '') {
                    // もしかして列追加前のキャッシュを見てる？いや getValues() は最新
                    // マイグレーション直後はデータが入ってないかも
                    // しかし列シフトで元の"金額"は列7に来ているはず
                }

                results.push({
                    date: Utilities.formatDate(d, timeZone, 'yyyy/MM/dd HH:mm'),
                    displayDate: row[0],
                    id: row[1],
                    name: row[2],
                    menuId: row[3],
                    price: row[7], // H列
                    phone: row[9], // J列
                    notes: row[10], // K列
                    status: row[11], // L列
                    googleEventId: row[12] || '' // M列
                });
            }
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
                if (data[i][1] == reservationId) {
                    // Update Column M (13th column)
                    sheet.getRange(i + 1, 13).setValue(eventId);
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
                        status: data[i][11], // L列
                        googleEventId: data[i][12] // M列
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
                    // 1. Update Status (L列 / Column 12)
                    sheet.getRange(i + 1, 12).setValue('キャンセル');

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
                        googleEventId: data[i][12], // M列
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

            var newDateObj = parseDatetimeString(newDatetime);
            if (!newDateObj) {
                return { success: false, message: '日時の形式が正しくありません: ' + newDatetime };
            }
            var newDatetimeStr = Utilities.formatDate(newDateObj, timeZone, 'yyyy/MM/dd HH:mm');

            if (currentDatetimeStr !== newDatetimeStr) {
                if (!this.reserveSlot(newDatetimeStr)) {
                    return { success: false, message: '変更先の日時は既に埋まっています。' };
                }
                this.updateSlotStatus(currentDatetimeStr, '空き');
            }

            var row = res.row;
            sheet.getRange(row, 1).setValue(formatDateJP(newDateObj));

            // メニュー更新
            if (newMenuId && newMenuId !== res.menu) {
                var menuItems = this.getMenuItems();
                var selectedMenu = menuItems.find(function (item) { return item.id === newMenuId; });
                if (selectedMenu) {
                    sheet.getRange(row, 4).setValue(newMenuId); // メニューID (D)
                    var price = parseInt(selectedMenu.price, 10) || 0;
                    var priceExcl = Math.floor(price / 1.1);
                    var tax = price - priceExcl;
                    sheet.getRange(row, 5).setValue(selectedMenu.name); // メニュー名 (E)
                    sheet.getRange(row, 6).setValue(priceExcl); // 税抜 (F)
                    sheet.getRange(row, 7).setValue(tax); // 税 (G)
                    sheet.getRange(row, 8).setValue(price); // 税込 (H)
                }
            } else if (!newMenuId) {
                // メニューが変わらない場合でも、価格列などが空なら埋めるべきだが、既存データの更新は今回は必須ではない
                // ただし、もしメニューIDしかない旧データに対して更新がかかった場合、ここでメニュー名を埋めてあげる親切設計はあり
                var currentResData = sheet.getRange(row, 1, 1, 8).getValues()[0];
                var currentMenuName = currentResData[4];
                if (!currentMenuName) {
                    var menuItems = this.getMenuItems();
                    var selectedMenu = menuItems.find(function (item) { return item.id === res.menu; });
                    if (selectedMenu) {
                        var price = parseInt(selectedMenu.price, 10) || 0;
                        var priceExcl = Math.floor(price / 1.1);
                        var tax = price - priceExcl;
                        sheet.getRange(row, 5).setValue(selectedMenu.name);
                        sheet.getRange(row, 6).setValue(priceExcl);
                        sheet.getRange(row, 7).setValue(tax);
                        sheet.getRange(row, 8).setValue(price);
                    }
                }
            }

            return {
                success: true,
                googleEventId: res.googleEventId,
                oldDate: res.date,
                newDate: newDateObj
            };
        }
    };

})();
