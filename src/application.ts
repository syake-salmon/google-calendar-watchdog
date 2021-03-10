declare var SpreadSheetsSQL: any;
const PROPERTY_KEY_SYNC_TOKEN: string = 'SYNC_TOKEN';
const PROPERTY_KEY_LINE_TOKEN: string = 'LINE_TOKEN';
const PROPERTY_KEY_SLACK_WEBHOOK_ENDPOINT = 'SLACK_WEBHOOK_ENDPOINT';
const FILE_ID_EVENTS: string = '1PVVkZUjD6wSw-kIGp1BpnLm_TsaYCWEtUEYEwE8QAaI';
const ENDPOINT_LINE_NOTIFY_API: string = 'https://notify-api.line.me/api/notify';

var properties: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties();
var nextSyncToken: string = '';

function calendarUpdated(event: GoogleAppsScript.Events.CalendarEventUpdated): void {
    console.time('----- calendarUpdated -----');

    try {
        var calendarId: string = event.calendarId;

        var options = {
            syncToken: getSyncToken(calendarId)
        };
        var recentlyUpdatedEvents: GoogleAppsScript.Calendar.Schema.Event[] = getRecentlyUpdatedEvents(calendarId, options);

        var message: string = generateMessage(recentlyUpdatedEvents);

        notifyLINE(message);

        refleshStoredEvents(calendarId);

        properties.setProperty(PROPERTY_KEY_SYNC_TOKEN, nextSyncToken);
    } catch (e) {
        console.error(e);
        var slackOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'post',
            payload: JSON.stringify({ 'username': 'google-calendar-watchdog', 'text': 'カレンダー変更通知処理中にエラーが発生しました。<https://script.google.com/home/projects/1VE5tPlGhiNWUOsJje9HYVOjX4BjvK-VLx5_8-LsV7A2StRMUsu3qXWuM/executions|[ログ]>\nERROR=>' + e.message }),
            muteHttpExceptions: true
        };
        callExternalAPI(properties.getProperty(PROPERTY_KEY_SLACK_WEBHOOK_ENDPOINT), slackOptions);
    }

    console.timeEnd('----- calendarUpdated -----');
}

function getSyncToken(calendarId: string): string {
    console.time('----- getSyncToken -----');

    var token: string = properties.getProperty(PROPERTY_KEY_SYNC_TOKEN);
    if (!token) {
        token = Calendar.Events.list(calendarId, { 'timeMin': (new Date()).toISOString() }).nextSyncToken;
    }

    console.timeEnd('----- getSyncToken -----');
    return token;
}

function getRecentlyUpdatedEvents(calendarId: string, options: object): GoogleAppsScript.Calendar.Schema.Event[] {
    console.time('----- getRecentlyUpdatedEvents -----');

    var events: GoogleAppsScript.Calendar.Schema.Events = Calendar.Events.list(calendarId, options);
    nextSyncToken = events.nextSyncToken;

    console.timeEnd('----- getRecentlyUpdatedEvents -----');
    return events.items;
}

function generateMessage(events: GoogleAppsScript.Calendar.Schema.Event[]): string {
    console.time('----- generateNotifyMessages -----');

    var message: string = '';
    var messages: string[] = [];
    for (var i: number = 0; i < events.length; i++) {
        var status: string = events[i].status;
        var storedEvent: StoredEvent = searchStoredEventById(events[i].id);
        if (status == 'cancelled') {
            if (storedEvent) {
                messages.push('Googleカレンダーの予定が削除されました。\n\nタイトル：' + storedEvent.summary + '\n開始：' + dateToString(storedEvent.start) + '\n終了：' + dateToString(storedEvent.end));
            } else {
                messages.push('Googleカレンダーの予定が削除されました。');
            }
        } else {
            var start: string = (events[i].start.dateTime) ? events[i].start.dateTime : events[i].start.date;
            var end: string = (events[i].end.dateTime) ? events[i].end.dateTime : events[i].end.date;

            if (storedEvent) {
                messages.push('Googleカレンダーの予定が更新されました。\n\nタイトル：' + events[i].summary + '\n開始：' + dateToString(start) + '\n終了：' + dateToString(end));
            } else {
                messages.push('Googleカレンダーに予定が登録されました。\n\nタイトル：' + events[i].summary + '\n開始：' + dateToString(start) + '\n終了：' + dateToString(end));
            }
        }
    }
    message = messages.join('\n----------\n');

    console.timeEnd('----- generateNotifyMessages -----');
    return message;
}

function searchStoredEventById(id: string): StoredEvent {
    console.time('----- searchStoredEventById -----');

    var event: StoredEvent = searchStoredEvents('id = ' + id)[0];

    console.timeEnd('----- searchStoredEventById -----');
    return event;
}

function searchStoredEvents(filter: string): StoredEvent[] {
    console.time('----- searchStoredEvents -----');

    var events: StoredEvent[] = [];
    var result: any[] = SpreadSheetsSQL.open(FILE_ID_EVENTS, 'DATA').select(['id', 'summary', 'start', 'end']).filter(filter).result();
    for (var i: number = 0; i < result.length; i++) {
        events.push(new StoredEvent(result[i].id, result[i].summary, result[i].start, result[i].end));
    }

    console.timeEnd('----- searchStoredEvents -----');
    return events;
}

function dateToString(source: string): string {
    console.time('----- dateToString -----');

    var stringFormat: string = '';
    var yyyyMMdd: string = String(source).split('T')[0];
    var hhmm: string = String(source).split('T')[1];

    var yyyy: string = String(yyyyMMdd).split('-')[0];
    var MM: string = String(yyyyMMdd).split('-')[1];
    var dd: string = String(yyyyMMdd).split('-')[2];
    var hh: string = (hhmm) ? String(hhmm).split(':')[0] : '00';
    var mm: string = (hhmm) ? String(hhmm).split(':')[1] : '00';
    stringFormat = yyyy + '-' + MM + '-' + dd + ' ' + hh + ':' + mm;

    console.timeEnd('----- dateToString -----');
    return stringFormat;
}

function notifyLINE(message: string): void {
    console.time('----- notifyLINE -----');

    var token: string = properties.getProperty(PROPERTY_KEY_LINE_TOKEN);
    var options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: 'post',
        payload: 'message=' + message,
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
    };
    callExternalAPI(ENDPOINT_LINE_NOTIFY_API, options);

    console.timeEnd('----- notifyLINE -----');
}

function refleshStoredEvents(calendarId: string): void {
    console.time('----- refleshStoredEvents -----');

    SpreadSheetsSQL.open(FILE_ID_EVENTS, 'DATA').deleteRows();
    var events: GoogleAppsScript.Calendar.Schema.Event[] = Calendar.Events.list(calendarId, { 'timeMin': (new Date()).toISOString() }).items;
    var storedEvents: StoredEvent[] = [];
    for (var i: number = 0; i < events.length; i++) {
        var start: string = (events[i].start.dateTime) ? events[i].start.dateTime : events[i].start.date;
        var end: string = (events[i].end.dateTime) ? events[i].end.dateTime : events[i].end.date;
        storedEvents.push(new StoredEvent(events[i].id, events[i].summary, start, end));
    }
    SpreadSheetsSQL.open(FILE_ID_EVENTS, 'DATA').insertRows(storedEvents);
    SpreadsheetApp.openById(FILE_ID_EVENTS).getSheetByName('DATA').getDataRange().setNumberFormat('@');

    console.timeEnd('----- refleshStoredEvents -----');
}

function callExternalAPI(endpoint: string, options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions) {
    console.time('----- callExternalAPI -----');

    var response = UrlFetchApp.fetch(endpoint, options);

    console.timeEnd('----- callExternalAPI -----');
    return response;
}

class StoredEvent {
    constructor(id: string, summary: string, start: string, end: string) {
        this.id = id;
        this.summary = summary;
        this.start = start;
        this.end = end;
    }

    id: string;
    summary: string;
    start: string;
    end: string;
}