const SHEET_URL = PropertiesService.getScriptProperties().getProperty('SHEET_URL');

interface NotificationType {
  Daily: 'daily';
  Weekly: 'weekly';
  Before: 'before';
}
class Notification {
  eid: string;
  url: string;

  constructor(public type: 'daily'|'weekly'|'before', public config: string, public slackChannelName: string, public calendarId: string, public eventId: string, public title: string, public start: GoogleAppsScript.Base.Date, public end: GoogleAppsScript.Base.Date) {
    const calendarIdFragment = this.calendarId.replace(/@group.calendar.google.com$/, '@g');
    const eventIdFragment = this.eventId.split('@')[0];
    this.eid = Utilities.base64Encode(`${eventIdFragment} ${calendarIdFragment}`).replace(/=$/, '');
    this.url = `https://calendar.google.com/calendar/u/0/event?eid=${this.eid}`;
  }
}

function syncCalendars() {
  const spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
  const calendarSheet = spreadsheet.getSheetByName('Calendars');
  const notificationSheet = spreadsheet.getSheetByName('Notifications');

  const rows : any[][] =  calendarSheet.getRange(2, 1, calendarSheet.getLastRow() - 1, 5).getValues();

  const notifications : Notification[] = [];
  
  rows.forEach((row) => {
    const calendarId = row[0];
    const slackChannelName = row[1];
    const dailyNotification = row[2];
    const weeklyNotification = row[3];
    const beforeNotification = row[4];

    const calendarEvents = fetchCalendarEvents(calendarId);

    if (dailyNotification && dailyNotification != '') {
      calendarEvents.forEach((event) => {
        notifications.push(new Notification('daily', dailyNotification.toString(), slackChannelName, calendarId, event.getId(), event.getTitle(), event.getStartTime(), event.getEndTime()));
      });
    }

    if (weeklyNotification && weeklyNotification != '') {
      calendarEvents.forEach((event) => {
        notifications.push(new Notification('weekly', weeklyNotification.toString(), slackChannelName, calendarId, event.getId(), event.getTitle(), event.getStartTime(), event.getEndTime()));
      });
    }

    if (beforeNotification && beforeNotification != '') {
      calendarEvents.forEach((event) => {
        notifications.push(new Notification('before', beforeNotification.toString(), slackChannelName, calendarId, event.getId(), event.getTitle(), event.getStartTime(), event.getEndTime()));
      });
    }
  });

  console.log(notifications);

  notificationSheet.clear();
  notificationSheet.appendRow(['Type', 'Config', 'Slack Channel', 'Calendar ID', 'Event ID', 'Title', 'Start', 'End', 'Eid', 'URL']);

  notifications.forEach((notification) => {
    const row = [
      notification.type,
      notification.config,
      notification.slackChannelName,
      notification.calendarId,
      notification.eventId,
      notification.title,
      notification.start,
      notification.end,
      notification.eid,
      notification.url,
    ];
    notificationSheet.appendRow(row);
  });
}

function fetchCalendarEvents(calendarId) : GoogleAppsScript.Calendar.CalendarEvent[] {
  const calendar = CalendarApp.getCalendarById(calendarId);
  const now = new Date();
  const nextWeek = new Date(now);
  nextWeek.setDate(nextWeek.getDate() + 7);
  return calendar.getEvents(now, nextWeek);
}

function notifyEvents() {
  const spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
  const eventSheet = spreadsheet.getSheetByName('Notifications');
}
