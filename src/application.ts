const PROPERTY_KEY_LINE_TOKEN: string = 'LINE_TOKEN';
const PROPERTY_KEY_SLACK_WEBHOOK_ENDPOINT = 'SLACK_WEBHOOK_ENDPOINT';
const ENDPOINT_LINE_NOTIFY_API: string = 'https://notify-api.line.me/api/notify';

var properties: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties();

interface CalendarEventUpdated {
    authMode: GoogleAppsScript.Script.AuthMode;
    calendarId: string;
    triggerUid: string;
}

/**
 * イベントが更新されたとき、対象のCalendarIdとEvent.summaryをコンソールに出力する。
 */
function onUpdatedEvent(event: GoogleAppsScript.Events.CalendarEventUpdated): void {
    console.time('----- onUpdatedEvent -----');
    console.log(`calendarId: ${event.calendarId}`);

    try {
        getUpdatedEvents(event.calendarId).forEach((e) => {
            let message: string;
            if (e.status === "cancelled") {
                message = `\nGoogleカレンダーの予定が削除されました。`
            } else {
                let startDateTime: string;
                let endDateTime: string;
                let location: string;

                startDateTime = new Date(e.start.dateTime).toLocaleString("ja-JP", { timeZone: e.start.timeZone });
                endDateTime = new Date(e.end.dateTime).toLocaleString("ja-JP", { timeZone: e.end.timeZone });
                location = e.location == undefined ? "" : e.location;

                message = `\nGoogleカレンダーの予定が更新されました。\n==========\nタイトル：${e.summary}\n開始日時：${startDateTime}\n終了日時：${endDateTime}\n場所      ：${location}`
            }

            notifyLINE(message);
        })
    } catch (e) {
        console.error(e);
        var slackOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'post',
            payload: JSON.stringify({ 'username': 'google-calendar-watchdog', 'text': 'カレンダー変更通知処理中にエラーが発生しました。<https://script.google.com/home/projects/1VE5tPlGhiNWUOsJje9HYVOjX4BjvK-VLx5_8-LsV7A2StRMUsu3qXWuM/executions|[ログ]>\nERROR=>' + e.message }),
            muteHttpExceptions: true
        };
        callExternalAPI(properties.getProperty(PROPERTY_KEY_SLACK_WEBHOOK_ENDPOINT), slackOptions);
    }

    console.timeEnd('----- onUpdatedEvent -----');
}

/**
 * @link https://developers.google.com/calendar/api/v3/reference/events/list?hl=ja#parameters
 */
class CalendarQueryOptions {
    maxResults?: number;
    syncToken?: string;
    timeMin?: string;
}

/**
 * 引数に指定されたCalendarIdにひもづくイベントのうち、前回取得した時より更新があったイベントのみ取得する。
 * 過去に一度もイベントを取得していない場合、過去30日分を取得する。
 *
 * @param {string} calendarId カレンダーを一意に識別するID
 * @returns {GoogleAppsScript.Calendar.Schema.Event[]} イベントの一覧（最大100件）
 */
function getUpdatedEvents(calendarId: string): GoogleAppsScript.Calendar.Schema.Event[] {
    console.time('----- getUpdatedEvents -----');

    const properties = PropertiesService.getUserProperties();
    const key = `syncToken: ${calendarId}`;
    const syncToken = properties.getProperty(key);

    let options: CalendarQueryOptions = { maxResults: 100 };
    if (syncToken) {
        options = { ...options, syncToken: syncToken };
    } else {
        options = { ...options, timeMin: getRelativeDate(-30, 0).toISOString() };
    }

    const events = Calendar.Events?.list(calendarId, options);

    if (events?.nextSyncToken) {
        properties.setProperty(key, events?.nextSyncToken);
    }

    console.timeEnd('----- getUpdatedEvents -----');
    return events?.items ? events.items : [];
}

/**
 * Helper function to get a new Date object relative to the current date.
 * @param {number} daysOffset The number of days in the future for the new date.
 * @param {number} hour The hour of the day for the new date, in the time zone of the script.
 * @return {Date} The new date.
 */
function getRelativeDate(daysOffset: number, hour: number): Date {
    var date = new Date();
    date.setDate(date.getDate() + daysOffset);
    date.setHours(hour);
    date.setMinutes(0);
    date.setSeconds(0);
    date.setMilliseconds(0);
    return date;
}

/**
 * 引数で指定されたメッセージをLINEに送信する。
 */
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

function callExternalAPI(endpoint: string, options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions) {
    console.time('----- callExternalAPI -----');

    var response = UrlFetchApp.fetch(endpoint, options);

    console.timeEnd('----- callExternalAPI -----');
    return response;
}
