function doPost(e) {
    var output = ContentService.createTextOutput();
    output.setMimeType(ContentService.MimeType.JSON);

    try {
        var data;
        if (e.postData && e.postData.contents) {
            data = JSON.parse(e.postData.contents);
        } else {
            throw new Error('No data received');
        }

        // --- Action Dispatcher ---
        var action = data.action;

        // --- Admin Actions ---
        if (action === 'login') {
            var props = PropertiesService.getScriptProperties();
            var validId = props.getProperty('ADMIN_ID');
            var validPass = props.getProperty('ADMIN_PASSWORD');

            if (data.id === validId && data.password === validPass) {
                output.setContent(JSON.stringify({ status: 'success', token: 'valid_session' })); // Simple token
            } else {
                output.setContent(JSON.stringify({ status: 'error', message: 'Authentication failed' }));
            }
            return output;
        }

        if (action === 'get_admin_data') {
            // Basic auth check (should be improved in production)
            // In a real app, validate token. Here we rely on the frontend having passed login.
            var menu = SheetUtils.getMenuItems();
            var slots = SheetUtils.getAvailableSlots();
            var reservations = SheetUtils.getReservations();
            output.setContent(JSON.stringify({ status: 'success', menu: menu, slots: slots, reservations: reservations }));
            return output;
        }

        if (action === 'update_slot_status') {
            if (SheetUtils.updateSlotStatus(data.datetime, data.status)) {
                output.setContent(JSON.stringify({ status: 'success' }));
            } else {
                output.setContent(JSON.stringify({ status: 'error', message: 'Slot not found' }));
            }
            return output;
        }

        if (action === 'update_slot_datetime') {
            var result = SheetUtils.updateSlotDatetime(data.currentDatetime, new Date(data.newDatetime));
            if (result.success) {
                output.setContent(JSON.stringify({ status: 'success' }));
            } else {
                output.setContent(JSON.stringify({ status: 'error', message: result.message }));
            }
            return output;
        }

        if (action === 'delete_slot') {
            if (SheetUtils.deleteSlot(data.datetime)) {
                output.setContent(JSON.stringify({ status: 'success' }));
            } else {
                output.setContent(JSON.stringify({ status: 'error', message: 'Slot not found' }));
            }
            return output;
        }

        if (action === 'save_menu') {
            SheetUtils.saveMenuItem(data.item);
            output.setContent(JSON.stringify({ status: 'success' }));
            return output;
        }

        if (action === 'add_slots') {
            try {
                SheetUtils.addSlots(data.slots);
                output.setContent(JSON.stringify({ status: 'success' }));
            } catch (e) {
                output.setContent(JSON.stringify({ status: 'error', message: e.toString() }));
            }
            return output;
        }

        if (action === 'delete_menu') {
            if (SheetUtils.deleteMenuItem(data.id)) {
                output.setContent(JSON.stringify({ status: 'success' }));
            } else {
                output.setContent(JSON.stringify({ status: 'error', message: 'Item not found' }));
            }
            return output;
        }

        if (action === 'reorder_menu') {
            SheetUtils.reorderMenuItems(data.ids);
            output.setContent(JSON.stringify({ status: 'success' }));
            return output;
        }

        if (action === 'cancel_reservation') {
            try {
                var res = SheetUtils.cancelReservation(data.reservationId);
                if (res.success) {
                    if (res.googleEventId) {
                        var props = PropertiesService.getScriptProperties();
                        var calendarId = props.getProperty('CALENDAR_ID');
                        if (calendarId) {
                            try {
                                var cal = CalendarApp.getCalendarById(calendarId);
                                var evt = cal.getEventById(res.googleEventId);
                                if (evt) evt.deleteEvent();
                            } catch (e) { console.error('Calendar Delete Error:', e); }
                        }
                    }
                    output.setContent(JSON.stringify({ status: 'success' }));
                } else {
                    output.setContent(JSON.stringify({ status: 'error', message: res.message }));
                }
            } catch (e) {
                output.setContent(JSON.stringify({ status: 'error', message: e.toString() }));
            }
            return output;
        }

        if (action === 'update_reservation') {
            try {
                // data: { reservationId, newDatetime, newMenuId }
                var res = SheetUtils.updateReservation(data.reservationId, data.newDatetime, data.newMenuId);
                if (res.success) {
                    if (res.googleEventId) {
                        var props = PropertiesService.getScriptProperties();
                        var calendarId = props.getProperty('CALENDAR_ID');
                        if (calendarId) {
                            try {
                                var menuItems = SheetUtils.getMenuItems();
                                // If menu changed, get new duration. If not, we might need to look up current menu?
                                // Simplified: If newMenuId provided, use it. If not, we might lack duration info 
                                // if we don't fetch the existing reservation's menu.
                                // SheetUtils.updateReservation handles the sheet update.
                                // Let's just assume for now we can get the duration easily if we know the menu.
                                // If newMenuId is null, use existing menu? Code.js doesn't know existing menu ID easily without query.
                                // Let's just fetch the reservation to be sure or trust the input?
                                // Better: SheetUtils already looked it up. Let's make SheetUtils return the menu ID used?
                                // For now, let's just use a default or fetch if passed.
                                // The user likely passes newMenuId if they want to change it.

                                var menuId = data.newMenuId;
                                if (!menuId) {
                                    // If not provided, we need to find what the current menu is to calculate end time.
                                    var currentRes = SheetUtils.getReservation(data.reservationId);
                                    if (currentRes) menuId = currentRes.menu;
                                }

                                var selectedMenu = menuItems.find(function (m) { return m.id === menuId; });
                                var duration = selectedMenu ? parseInt(selectedMenu.duration, 10) : 60;

                                var newStartTime = new Date(data.newDatetime);
                                var newEndTime = new Date(newStartTime.getTime() + duration * 60000);

                                var cal = CalendarApp.getCalendarById(calendarId);
                                var evt = cal.getEventById(res.googleEventId);
                                if (evt) {
                                    evt.setTime(newStartTime, newEndTime);
                                    if (selectedMenu) {
                                        // Update description/title if menu changed? 
                                        // For simplicity, maybe just time for now unless requested.
                                        // But keeping it synced is better.
                                        var currentDesc = evt.getDescription();
                                        // Replacing menu name in description is complex without parsing.
                                        // Let's just update time.
                                    }
                                }
                            } catch (e) { console.error('Calendar Update Error:', e); }
                        }
                    }
                    output.setContent(JSON.stringify({ status: 'success' }));
                } else {
                    output.setContent(JSON.stringify({ status: 'error', message: res.message }));
                }
            } catch (e) {
                output.setContent(JSON.stringify({ status: 'error', message: e.toString() }));
            }
            return output;
        }


        // If 'action' is specified, handle specific tasks (like refreshing cache)
        if (action === 'refresh_menu') {
            var menu = SheetUtils.updateMenuCache();
            output.setContent(JSON.stringify({
                status: 'success',
                message: 'Menu cache updated',
                menu: menu
            }));
            return output;
        }

        // 1. Honeypot check
        if (data.honeypot) {
            console.warn('Bot detected via honeypot');
            output.setContent(JSON.stringify({ status: 'success', message: 'Received' }));
            return output;
        }

        // 2. Available Slot Check (Exclusive Lock)
        // Attempt to update the slot status. If it fails (already taken), return error.
        if (!SheetUtils.reserveSlot(data.datetime)) {
            throw new Error('選択された日時は既に予約が入ってしまいました。別の日時を選択してください。');
        }

        // 3. Input Validation
        if (!data.menu || !data.datetime || !data.name || !data.phone) {
            throw new Error('必須項目が不足しています');
        }
        if (data.name.length > 50) throw new Error('名前が長すぎます');
        if (data.phone.length > 20) throw new Error('電話番号の形式が正しくありません');
        if (data.notes && data.notes.length > 500) throw new Error('備考欄は500文字以内で入力してください');

        // -----------------------

        // Save to spreadsheet
        var id = SheetUtils.appendReservation(data);

        // --- Notifications ---
        var eventId = null;
        try {
            eventId = sendAdminNotifications(data, id);
        } catch (e) {
            console.error('Notification Error:', e);
            // Don't fail the request just because notification failed, but log it.
        }

        if (eventId) {
            SheetUtils.setReservationGoogleEventId(id, eventId);
        }

        // --- LINE Notification to User ---
        if (data.userId) {
            try {
                sendLineNotification(data.userId, data);
            } catch (e) {
                console.error('LINE Notification Error:', e);
            }
        }

        var result = {
            status: 'success',
            message: '予約を受け付けました',
            reservationId: id
        };
        output.setContent(JSON.stringify(result));

    } catch (error) {
        var errorResult = {
            status: 'error',
            message: 'Error: ' + error.toString()
        };
        output.setContent(JSON.stringify(errorResult));
    }

    return output;
}

function doGet(e) {
    var output = ContentService.createTextOutput();
    output.setMimeType(ContentService.MimeType.JSON);

    // Combine data fetching to single endpoint for performance
    if (e.parameter && e.parameter.action === 'get_data') {
        try {
            var menuItems = SheetUtils.getMenuItems();
            var slots = SheetUtils.getAvailableSlots();

            output.setContent(JSON.stringify({
                status: 'success',
                menu: menuItems,
                slots: slots
            }));
        } catch (err) {
            output.setContent(JSON.stringify({
                status: 'error',
                message: err.toString()
            }));
        }
        return output;
    }

    // Legacy or single fetch support if needed
    // For simplicity, let's stick to get_data as the main one for the new frontend
    // if (e.parameter && e.parameter.action === 'get_menu') {
    //     try {
    //         var menuItems = SheetUtils.getMenuItems();
    //         output.setContent(JSON.stringify({
    //             status: 'success',
    //             menu: menuItems
    //         }));
    //     } catch (err) {
    //         output.setContent(JSON.stringify({
    //             status: 'error',
    //             message: err.toString()
    //         }));
    //     }
    //     return output;
    // }

    // Default response
    output.setContent(JSON.stringify({ status: 'running', message: 'GAS Backend is active. Use ?action=get_data' }));
    return output;
}

// --- Trigger Setup ---

/**
 * Run this function ONCE from the GAS Editor to install the trigger.
 */
function setupTriggers() {
    // Prevent duplicate triggers
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'onSpreadsheetEdit') {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ScriptApp.newTrigger('onSpreadsheetEdit')
        .forSpreadsheet(ss)
        .onEdit()
        .create();

    console.log('Trigger set up successfully.');
}

/**
 * Triggered automatically when the spreadsheet is edited.
 * Checks if the edited sheet is 'メニュー設定', and if so, refreshes the cache.
 */
function onSpreadsheetEdit(e) {
    if (!e) return;

    var sheetName = e.range.getSheet().getName();

    if (sheetName === 'メニュー設定') {
        console.log('Menu sheet edited. Updating cache...');
        SheetUtils.updateMenuCache();
    }
}

// --- Sidebar & Menus ---

function onOpen(e) {
    SpreadsheetApp.getUi()
        .createMenu('予約管理')
        .addItem('予約枠設定', 'showSidebar')
        .addToUi();
}

function showSidebar() {
    var html = HtmlService.createHtmlOutputFromFile('sidebar')
        .setTitle('予約枠登録')
        .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

function addSlotsFromSidebar(slots) {
    try {
        SheetUtils.addSlots(slots);
        return { status: 'success' };
    } catch (error) {
        throw new Error(error.toString());
    }
}

function sendAdminNotifications(data, reservationId) {
    var props = PropertiesService.getScriptProperties();
    var adminEmail = props.getProperty('ADMIN_EMAIL');
    var calendarId = props.getProperty('CALENDAR_ID');

    // 1. Get Menu Details for Duration & Name
    var menuItems = SheetUtils.getMenuItems();
    var selectedMenu = menuItems.find(function (item) { return item.id === data.menu; });
    var menuName = selectedMenu ? selectedMenu.name : '不明なメニュー';
    var duration = selectedMenu ? parseInt(selectedMenu.duration, 10) : 60; // Default 60 min

    var startTime = new Date(data.datetime);
    var endTime = new Date(startTime.getTime() + duration * 60000);

    // 2. Send Email (if ADMIN_EMAIL is set)
    if (adminEmail) {
        var subject = '【予約受信】' + data.name + '様 (' + menuName + ')';
        var body =
            '新しい予約が入りました。\n\n' +
            '予約ID: ' + reservationId + '\n' +
            '日時: ' + Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') + '\n' +
            'お名前: ' + data.name + '\n' +
            '電話番号: ' + data.phone + '\n' +
            'メニュー: ' + menuName + ' (' + duration + '分)\n' +
            '備考: ' + (data.notes || 'なし') + '\n\n' +
            '管理者アプリで確認してください。';
        GmailApp.sendEmail(adminEmail, subject, body);
    } else {
        console.warn('ADMIN_EMAIL not set. Skipping email notification.');
    }

    // 3. Add to Google Calendar (if CALENDAR_ID is set)
    var eventId = null;
    if (calendarId) {
        try {
            var cal = CalendarApp.getCalendarById(calendarId);
            if (!cal) {
                console.warn('Could not find calendar for ' + calendarId + '. Skipping event creation.');
            } else {
                var event = cal.createEvent('予約: ' + data.name + '様', startTime, endTime, {
                    description: 'メニュー: ' + menuName + '\n電話: ' + data.phone + '\n備考: ' + data.notes
                });
                eventId = event.getId();
            }
        } catch (e) {
            console.error('Calendar Error:', e);
        }
    } else {
        console.warn('CALENDAR_ID not set. Skipping calendar event.');
    }
    return eventId;
}

/**
 * Run this function from the GAS Editor to authorize scopes.
 */
function authorize() {
    // Access services to trigger auth dialog
    GmailApp.getDrafts();
    CalendarApp.getDefaultCalendar();
    console.log('Authorization successful');
}

/**
 * Sends a push message to the specified LINE User ID.
 */
function sendLineNotification(userId, data) {
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');

    if (!token) {
        console.warn('LINE_CHANNEL_ACCESS_TOKEN not set. Skipping LINE notification.');
        return;
    }

    var menuItems = SheetUtils.getMenuItems();
    var selectedMenu = menuItems.find(function (item) { return item.id === data.menu; });
    var menuName = selectedMenu ? selectedMenu.name : '不明なメニュー (ID: ' + data.menu + ')';

    // Format Date
    var d = new Date(data.datetime);
    var dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'M月d日(E) HH:mm');
    // Japanese day of week manual map if locale not reliable, but 'E' usually works
    // Let's force Japanese day of week to be safe
    var days = ['日', '月', '火', '水', '木', '金', '土'];
    var dayName = days[d.getDay()];
    var dateStrJP = Utilities.formatDate(d, Session.getScriptTimeZone(), 'M月d日') + '(' + dayName + ') ' + Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');

    var messageText =
        data.name + '様、ご予約ありがとうございます。\n' +
        '以下の内容で承りました。\n\n' +
        '■日時: ' + dateStrJP + '\n' +
        '■メニュー: ' + menuName + '\n' +
        '■金額: ' + Number(selectedMenu ? selectedMenu.price : 0).toLocaleString() + '円\n\n' +
        'ご来店をお待ちしております。';

    var payload = {
        to: userId,
        messages: [{
            type: 'text',
            text: messageText
        }]
    };

    var options = {
        method: 'post',
        contentType: 'application/json',
        headers: {
            'Authorization': 'Bearer ' + token
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try {
        var response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
        var rCode = response.getResponseCode();
        var rText = response.getContentText();
        console.log('LINE API Response: ' + rCode + ' / ' + rText);
    } catch (e) {
        console.error('Failed to send LINE message: ' + e.toString());
    }
}
