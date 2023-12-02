const PROPERTY_KEY_LINE_TOKEN: string = 'LINE_TOKEN';
const PROPERTY_KEY_SLACK_WEBHOOK_ENDPOINT = 'SLACK_WEBHOOK_ENDPOINT';
const ENDPOINT_LINE_NOTIFY_API: string = 'https://notify-api.line.me/api/notify';
const PROPERTIES: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties();

interface CalendarEventUpdated {
    authMode: GoogleAppsScript.Script.AuthMode;
    calendarId: string;
    triggerUid: string;
}

/**
 * Googleカレンダーのイベントが更新されたとき、更新内容をLINE Notifier APIへ通知する。
 *
 * @param {GoogleAppsScript.Events.CalendarEventUpdated} event Googleカレンダー更新トリガから通知されるイベント情報
 * @throws {Exception} 通知処理内で発生した例外。コンソールおよび保守用Slackチャンネルへ通知する。
 */
function onUpdatedEvent(event: GoogleAppsScript.Events.CalendarEventUpdated): void {
    console.time('onUpdatedEvent');
    console.log(`Updated Calendar. calendarId: ${event.calendarId}`);

    try {
        getUpdatedEvents(event.calendarId).forEach((e) => {
            let message: string;
            if (e.status === "cancelled") {
                message = `\nGoogleカレンダーの予定が削除されました。\n==========\nタイトル：${e.summary}`
            } else {
                let startDateTime: string;
                let endDateTime: string;
                let location: string;

                // Eventの開始・終了日時（UST）をEventに設定されたタイムゾーンに変換する
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

        callExternalAPI(PROPERTIES.getProperty(PROPERTY_KEY_SLACK_WEBHOOK_ENDPOINT), slackOptions);
        throw e;
    }

    console.timeEnd('onUpdatedEvent');
}

/**
 * @link https://developers.google.com/calendar/api/v3/reference/events/list?hl=ja#parameters
 */
class CalendarQueryOptions {
    maxResults?: number;
    showDeleted?: boolean;
    syncToken?: string;
    timeMin?: string;
}

/**
 * 引数に指定されたCalendarIdにひもづくイベントのうち、前回取得した時より更新があったイベントのみ取得する。
 * syncTokenが取得できなかった場合や過去に一度もイベントを取得していない場合、過去10日分を取得する。
 *
 * @param {string} calendarId カレンダーを一意に識別するID
 * @returns {GoogleAppsScript.Calendar.Schema.Event[]} イベントの一覧（最大100件）
 */
function getUpdatedEvents(calendarId: string): GoogleAppsScript.Calendar.Schema.Event[] {
    console.time('getUpdatedEvents');

    const key = `syncToken: ${calendarId}`;
    const syncToken = PROPERTIES.getProperty(key);

    let options: CalendarQueryOptions = { maxResults: 100, showDeleted: true };
    if (syncToken) {
        options = { ...options, syncToken: syncToken };
    } else {
        options = { ...options, timeMin: getRelativeDate(-10, 0).toISOString() };
    }

    const events = Calendar.Events?.list(calendarId, options);

    if (events?.nextSyncToken) {
        PROPERTIES.setProperty(key, events?.nextSyncToken);
    }

    console.timeEnd('getUpdatedEvents');
    return events?.items ? events.items : [];
}

/**
 * 引数で指定されたメッセージをLINE Notifier APIに送信する。
 *
 * @param {string} message 通知メッセージ。
 */
function notifyLINE(message: string): void {
    console.time('notifyLINE');

    var token: string = PROPERTIES.getProperty(PROPERTY_KEY_LINE_TOKEN);
    var options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: 'post',
        payload: 'message=' + message,
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
    };
    callExternalAPI(ENDPOINT_LINE_NOTIFY_API, options);

    console.timeEnd('notifyLINE');
}

/**
 * 現在日付に引数で指定されたオフセットを加算した日付のDateオブジェクトを取得するユーティリティ関数。
 *
 * @param {number} daysOffset 現在日に加算する日数。
 * @param {number} hour 設定する時刻。
 * @return {Date} 
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
 * REST API呼び出しを行なうユーティリティ関数。
 *
 * @param {string} endpoint 呼び出すAPIのエンドポイントURL。
 * @param {GoogleAppsScript.URL_Fetch.URLFetchRequestOptions} options リクエストヘッダに付与するオプション。
 */
function callExternalAPI(endpoint: string, options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions): GoogleAppsScript.URL_Fetch.HTTPResponse {
    var response = UrlFetchApp.fetch(endpoint, options);
    return response;
}
