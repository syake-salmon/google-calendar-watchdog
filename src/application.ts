declare var SpreadSheetsSQL: any;
const PROPERTY_KEY_SYNC_TOKEN: string = 'SYNC_TOKEN';
const PROPERTY_KEY_LINE_TOKEN: string = 'LINE_TOKEN';
const FILE_ID_EVENTS: string = '1PVVkZUjD6wSw-kIGp1BpnLm_TsaYCWEtUEYEwE8QAaI';
var properties: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties();

function calendarUpdated(event: GoogleAppsScript.Events.CalendarEventUpdated): void {
    console.log('calendarUpdated >>>>>>>>>>');

    var calendarId: string = event.calendarId;

    var options = {
        syncToken: properties.getProperty(PROPERTY_KEY_SYNC_TOKEN)
    };

    if (options.syncToken) {
        console.info('syncToken from property=[%s]', options.syncToken);
    } else {
        options.syncToken = Calendar.Events.list(calendarId, { 'timeMin': (new Date()).toISOString() }).nextSyncToken;
        console.info('syncToken from calendar=[%s]', options.syncToken);
    }

    var updatedEvents: GoogleAppsScript.Calendar.Schema.Events = Calendar.Events.list(calendarId, options);
    for (var i: number = 0; i < updatedEvents.items.length; i++) {
        var eventId: string = updatedEvents.items[i].id;
        var status: string = updatedEvents.items[i].status;
        var summary: string = '';
        var start: GoogleAppsScript.Calendar.Schema.EventDateTime;
        var end: GoogleAppsScript.Calendar.Schema.EventDateTime;
        var message: string = '';

        var rawResult: any[] = SpreadSheetsSQL.open(FILE_ID_EVENTS, 'DATA').select(['id', 'summary', 'start', 'end']).filter('id = ' + eventId).result();
        if (status == 'cancelled') {
            summary = rawResult[0].summary;
            start = rawResult[0].start;
            end = rawResult[0].end;
            message = 'Googleカレンダーの予定が削除されました。\n\nタイトル：' + summary + '\n開始：' + start + '\n終了：' + end;
        } else {
            if (rawResult.length == 0) {
                summary = updatedEvents.items[i].summary;
                start = updatedEvents.items[i].start;
                end = updatedEvents.items[i].end;
                message = 'Googleカレンダーに予定が登録されました。\n\nタイトル：' + summary + '\n開始：' + start + '\n終了：' + end;
            } else {
                summary = updatedEvents.items[i].summary;
                start = updatedEvents.items[i].start;
                end = updatedEvents.items[i].end;
                message = 'Googleカレンダーの予定が更新されました。\n\nタイトル：' + summary + '\n開始：' + start + '\n終了：' + end;
            }
        }
        
        var token: string = properties.getProperty(PROPERTY_KEY_LINE_TOKEN);
        var lineOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'post',
            payload: 'message=' + message,
            headers: {'Authorization' : 'Bearer ' + token},
            muteHttpExceptions: true
        };
        UrlFetchApp.fetch('https://notify-api.line.me/api/notify', lineOptions)
    }

    SpreadSheetsSQL.open(FILE_ID_EVENTS, 'DATA').deleteRows();
    var latestEventsArr: GoogleAppsScript.Calendar.Schema.Event[] = Calendar.Events.list(calendarId).items;
    for (var i: number = 0; i < latestEventsArr.length; i++) {
        SpreadSheetsSQL.open(FILE_ID_EVENTS, 'DATA').insertRows([
            { id: latestEventsArr[i].id, summary: latestEventsArr[i].summary, start: latestEventsArr[i].start, end: latestEventsArr[i].end }
        ]);
    }

    properties.setProperty(PROPERTY_KEY_SYNC_TOKEN, updatedEvents.nextSyncToken);
    console.log('saved next syncToken=[%s]', updatedEvents.nextSyncToken);

    console.log('calendarUpdated <<<<<<<<<<');
}
