function doPost(e) {
    var output = ContentService.createTextOutput();
    output.setMimeType(ContentService.MimeType.JSON);

    try {
        // Parse the POST data
        // When sending JSON from fetch, it might come as postData.contents
        var data;
        if (e.postData && e.postData.contents) {
            data = JSON.parse(e.postData.contents);
        } else {
            throw new Error('No data received');
        }

        // Save to spreadsheet
        var id = SheetUtils.appendReservation(data);

        // Return success response with CORS headers simulated by text output
        // Note: GAS Web App redirects make true CORS headers complex, 
        // but returning JSON usually works if client follows redirects.
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

// Keep doGet distinct for debug or simple check
function doGet(e) {
    var output = ContentService.createTextOutput(JSON.stringify({ status: 'running', message: 'GAS Backend is active' }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
}
