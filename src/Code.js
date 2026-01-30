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
            output.setContent(JSON.stringify({ status: 'success', menu: menu, slots: slots }));
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

        if (action === 'delete_menu') {
            if (SheetUtils.deleteMenuItem(data.id)) {
                output.setContent(JSON.stringify({ status: 'success' }));
            } else {
                output.setContent(JSON.stringify({ status: 'error', message: 'Item not found' }));
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
