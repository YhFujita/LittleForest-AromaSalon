var SheetUtils = (function () {
    // Updated: Fix Date Parsing & Add Columns
    // var SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'); 

    var SHEET_NAME_RESERVATIONS = '予約一覧';
    var SHEET_NAME_MENU = 'メニュー設定';
    var SHEET_NAME_SLOTS = '予約可能日時';
    var SHEET_NAME_SETTINGS = '設定';
    var CACHE_KEY_MENU = 'MENU_CACHE';
    var CACHE_KEY_SLOTS = 'SLOTS_CACHE';
    var MENU_CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes
    var SLOTS_CACHE_TTL_MS = 0; // Always refresh slots to prioritize consistency over cache hit

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
                sheet.appendRow(['ID', 'メニュー名', '価格', '所要時間(分)', '説明', '表示順', 'カテゴリタイトル', 'セクション説明', 'オプション']);
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
        } else if (name === SHEET_NAME_RESERVATIONS) {
            // Check for Migration (Add MenuName, TaxExcl, Tax Columns)
            // Old Header: [希望日時, 予約ID, 予約者名, メニュー, 金額, ...]
            // New Header: [希望日時, 予約ID, 予約者名, メニュー(ID), メニュー名, 税抜金額, 消費税, 金額(税込), ...]
            var header = sheet.getRange(1, 1, 1, 10).getValues()[0];
            // E列(index 4)が「金額」だったら旧形式
            if (header[4] === '金額') {
                // D列(index 3, メニュー)の後に3列挿入
                sheet.insertColumnsAfter(4, 3);
                // ヘッダー更新
                sheet.getRange(1, 4).setValue('メニュー(ID)');
                sheet.getRange(1, 5).setValue('メニュー名');
                sheet.getRange(1, 6).setValue('税抜金額');
                sheet.getRange(1, 7).setValue('消費税');
                sheet.getRange(1, 8).setValue('金額(税込)');
                // 以降のヘッダーは自動的にずれているはずだが、念のため確認
                // F->I, G->J, H->K, I->L, J->M
            }
        } else if (name === SHEET_NAME_MENU) {
            // Check for Migration (Add SectionDesc Column)
            if (sheet.getLastColumn() < 8) {
                sheet.getRange(1, 8).setValue('セクション説明');
            }
            if (sheet.getLastColumn() < 9) {
                sheet.getRange(1, 9).setValue('オプション');
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

    // 日付文字列のパース用ヘルパー関数
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

    function readCache(key) {
        var raw = PropertiesService.getScriptProperties().getProperty(key);
        if (!raw) return null;

        try {
            var parsed = JSON.parse(raw);
            // New format: { data: [...], updatedAt: number }
            if (parsed && typeof parsed === 'object' && parsed.data !== undefined && parsed.updatedAt !== undefined) {
                return parsed;
            }
            // Legacy format: cache body itself
            return {
                data: parsed,
                updatedAt: 0
            };
        } catch (e) {
            console.warn('Cache parse failed for key=' + key + '. Rebuilding cache.', e);
            return null;
        }
    }

    function writeCache(key, data) {
        var payload = {
            updatedAt: Date.now(),
            data: data
        };
        PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(payload));
    }

    function isCacheFresh(cache, ttlMs) {
        if (!cache) return false;
        if (ttlMs <= 0) return false;
        if (!cache.updatedAt || cache.updatedAt <= 0) return false;
        return (Date.now() - cache.updatedAt) <= ttlMs;
    }

    return {
        appendReservation: function (data) {
            var sheet = getSheet(SHEET_NAME_RESERVATIONS);
            var id = Utilities.getUuid();
            var timestamp = new Date();
            var formattedTimestamp = formatDateJP(timestamp);
            var bookingDate = new Date(data.datetime);
            var formattedBookingDate = formatDateJP(bookingDate);

            // Lookup Price & Menu Name
            var menuItems = this.getMenuItems();
            var selectedMenu = menuItems.find(function (item) { return item.id === data.menu; });
            var totalPrice = 0;
            var menuNames = [];

            if (selectedMenu) {
                totalPrice += parseInt(selectedMenu.price, 10) || 0;
                menuNames.push(selectedMenu.name);
            }

            // Sum options
            if (data.options && data.options.length > 0) {
                data.options.forEach(function (optId) {
                    var opt = menuItems.find(function (item) { return item.id === optId; });
                    if (opt) {
                        totalPrice += parseInt(opt.price, 10) || 0;
                        menuNames.push(opt.name);
                    }
                });
            }

            var finalMenuName = menuNames.join(' + ');

            // Calculate Tax (totalPrice is Incl Tax)
            var taxRate = 0.10;
            var priceExcl = Math.floor(totalPrice / 1.1);
            var tax = totalPrice - priceExcl;

            sheet.appendRow([
                formattedBookingDate, // A
                id,                   // B
                data.name,            // C
                data.menu,            // D (Primary Menu ID)
                finalMenuName,        // E (Menu Name(s))
                priceExcl,            // F (Excl Tax)
                tax,                  // G (Tax)
                totalPrice,           // H (Incl Tax)
                formattedTimestamp,   // I
                "'" + data.phone,     // J
                data.notes,           // K
                '受付',               // L
                data.googleEventId || '' // M
            ]);
            return id;
        },

        getMenuItems: function () {
            var cachedMenu = readCache(CACHE_KEY_MENU);
            if (cachedMenu && isCacheFresh(cachedMenu, MENU_CACHE_TTL_MS)) {
                return cachedMenu.data;
            }
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
                    section: row[6] || '',
                    sectionDesc: row[7] || '',
                    isOption: !!row[8]
                };
            }).filter(function (item) { return item.id && item.name; })
                .sort(function (a, b) {
                    var orderA = parseInt(a.order) || 9999;
                    var orderB = parseInt(b.order) || 9999;
                    return orderA - orderB;
                });

            writeCache(CACHE_KEY_MENU, menuItems);
            return menuItems;
        },

        getAvailableSlots: function () {
            var cachedSlots = readCache(CACHE_KEY_SLOTS);
            if (cachedSlots && isCacheFresh(cachedSlots, SLOTS_CACHE_TTL_MS)) {
                return cachedSlots.data;
            }
            return this.updateSlotsCache();
        },

        updateSlotsCache: function () {
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var data = sheet.getDataRange().getValues();
            var headers = data.shift();

            var slots = [];
            data.forEach(function (row) {
                if (row[1] !== '空き') return;
                var date = new Date(row[0]);
                if (isNaN(date.getTime())) {
                    // Ignore malformed rows so one bad row does not poison cache refresh.
                    return;
                }
                slots.push({
                    value: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
                    display: formatDateJP(date)
                });
            });

            slots.sort(function (a, b) {
                return a.value.localeCompare(b.value);
            });

            writeCache(CACHE_KEY_SLOTS, slots);
            return slots;
        },

        reserveSlot: function (targetDatetimeStr, durationMinutes) {
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var data = sheet.getDataRange().getValues();
            
            var startObj = parseDatetimeString(targetDatetimeStr);
            if (!startObj) return false;
            var endObj = new Date(startObj.getTime() + (durationMinutes || 60) * 60000);
            
            var rowsToUpdate = [];

            for (var i = 1; i < data.length; i++) {
                var cellDate = new Date(data[i][0]);
                if (isNaN(cellDate.getTime())) continue;
                
                if (cellDate >= startObj && cellDate < endObj) {
                    var cellStatus = data[i][1];
                    if (cellStatus !== '空き') {
                        return false;
                    }
                    rowsToUpdate.push(i + 1);
                }
            }
            
            if (rowsToUpdate.length === 0) return false;

            rowsToUpdate.forEach(function(row) {
                sheet.getRange(row, 2).setValue('予約済');
            });
            
            this.updateSlotsCache();
            return true;
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
            console.log('saveMenuItem called with:', JSON.stringify(item));
            var sheet = getSheet(SHEET_NAME_MENU);
            var data = sheet.getDataRange().getValues();
            var updated = false;

            if (item.id) {
                for (var i = 1; i < data.length; i++) {
                    if (data[i][0] == item.id) {
                        var range = sheet.getRange(i + 1, 1, 1, 9);
                        range.setValues([[
                            item.id,
                            item.name,
                            item.price,
                            item.duration,
                            item.description,
                            item.order,
                            item.section,
                            item.sectionDesc,
                            item.isOption ? 'TRUE' : ''
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
                    item.section,
                    item.sectionDesc,
                    item.isOption ? 'TRUE' : ''
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
                        menuName: data[i][4],
                        status: data[i][11], // L列
                        googleEventId: data[i][12] // M列
                    };
                }
            }
            return null;
        },

        releaseSlots: function(dateObj, menuId, menuNameStr) {
            var menuItems = this.getMenuItems();
            var duration = 60; // base if not found
            var primaryMenu = menuItems.find(function(item) { return item.id === menuId; });
            if (primaryMenu) {
                duration = parseInt(primaryMenu.duration, 10) || 0;
            }
            if (menuNameStr) {
                menuItems.forEach(function(item) {
                    if (item.isOption && menuNameStr.indexOf(item.name) !== -1) {
                        duration += parseInt(item.duration, 10) || 0;
                    }
                });
            }
            duration += 60; // buffer
            
            var endObj = new Date(dateObj.getTime() + duration * 60000);
            var sheet = getSheet(SHEET_NAME_SLOTS);
            var data = sheet.getDataRange().getValues();
            var rowsToUpdate = [];
            for (var i = 1; i < data.length; i++) {
                var cellDate = new Date(data[i][0]);
                if (isNaN(cellDate.getTime())) continue;
                if (cellDate >= dateObj && cellDate < endObj) {
                    rowsToUpdate.push(i + 1);
                }
            }
            if (rowsToUpdate.length > 0) {
                rowsToUpdate.forEach(function(row) {
                    sheet.getRange(row, 2).setValue('空き');
                });
                this.updateSlotsCache();
            }
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
                    this.releaseSlots(date, data[i][3], data[i][4]);

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

        updateReservation: function (reservationId, newDatetime, newMenuId, durationMinutes) {
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

            var menuItems = this.getMenuItems();
            var oldDuration = 60;
            var oldPrimaryMenu = menuItems.find(function(item) { return item.id === res.menu; });
            if (oldPrimaryMenu) oldDuration = parseInt(oldPrimaryMenu.duration, 10) || 0;
            if (res.menuName) {
                menuItems.forEach(function(item) {
                    if (item.isOption && res.menuName.indexOf(item.name) !== -1) oldDuration += parseInt(item.duration, 10) || 0;
                });
            }
            oldDuration += 60;

            if (currentDatetimeStr !== newDatetimeStr || oldDuration !== durationMinutes) {
                // Rollback approach:
                this.releaseSlots(res.date, res.menu, res.menuName);
                if (!this.reserveSlot(newDatetimeStr, durationMinutes)) {
                    // Rollback
                    this.reserveSlot(currentDatetimeStr, oldDuration);
                    return { success: false, message: '変更先の日時は既に埋まっているか、必要な連続枠がありません。' };
                }
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
