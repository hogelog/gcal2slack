const SHEET_URL = PropertiesService.getScriptProperties().getProperty('SHEET_URL');
const SLACK_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN')

const TICK_MINUTES = 5;

interface NotificationType {
  Daily: 'daily';
  Weekly: 'weekly';
  Before: 'before';
}
class Notification {
  eid: string;
  url: string;

  constructor(public configId: string, public type: 'daily'|'weekly'|'before', public config: string, public slackChannelName: string, public calendarId: string, public calendarTitle: string, public eventId: string, public eventTitle: string, public start: GoogleAppsScript.Base.Date, public end: GoogleAppsScript.Base.Date) {
    const calendarIdFragment = this.calendarId.replace(/@group.calendar.google.com$/, '@g');
    const eventIdFragment = this.eventId.split('@')[0];
    this.eid = Utilities.base64Encode(`${eventIdFragment} ${calendarIdFragment}`).replace(/=$/, '');
    this.url = `https://calendar.google.com/calendar/event?eid=${this.eid}`;
  }
}

class EventTime {
  date: Date;

  constructor(date: Date|GoogleAppsScript.Base.Date) {
    this.date = date instanceof Date ? new Date(date) : new Date(date.getTime());
  }

  dup(): EventTime {
    return new EventTime(new Date(this.date));
  }

  addDays(days: number): EventTime {
    const copy = this.dup();
    copy.date.setDate(copy.date.getDate() + days);
    return copy;
  }

  addMinutes(minutes: number): EventTime {
    const copy = this.dup();
    copy.date.setMinutes(copy.date.getMinutes() + minutes);
    return copy;
  }
}

class EventTimeRange {
  constructor(public start: EventTime, public end: EventTime) {
  }

  includes(time: EventTime): boolean {
    return this.start.date <= time.date && time.date < this.end.date;
  }
}

function syncCalendars() {
  const spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
  const calendarSheet = spreadsheet.getSheetByName('Calendars');
  const notificationSheet = spreadsheet.getSheetByName('Notifications');

  const rows : any[][] =  calendarSheet.getRange(2, 1, calendarSheet.getLastRow() - 1, 5).getValues();

  const notifications : Notification[] = [];
  const now = new EventTime(new Date());
  const tomorrow = now.addDays(1);
  const nextWeek = now.addDays(7);
  
  rows.forEach((row) => {
    const configId = Utilities.getUuid();
    const calendarId = row[0];
    const slackChannelName = row[1];
    const dailyNotification = row[2];
    const weeklyNotification = row[3];
    const beforeNotification = row[4];

    const calendar = CalendarApp.getCalendarById(calendarId);
    const calendarEvents = calendar.getEvents(now.date, nextWeek.date);

    if (dailyNotification && dailyNotification != '') {
      calendarEvents.forEach((event) => {
        const start = event.getStartTime();
        if (start < now.date || start.getDate() != now.date.getDate()) return;

        notifications.push(new Notification(configId, 'daily', dailyNotification.toString(), slackChannelName, calendarId, calendar.getName(), event.getId(), event.getTitle(), event.getStartTime(), event.getEndTime()));
      });
    }

    if (weeklyNotification && weeklyNotification != '') {
      calendarEvents.forEach((event) => {
        const start = event.getStartTime();
        if (start < now.date || start.getDay() < now.date.getDay()) return;

        notifications.push(new Notification(configId, 'weekly', weeklyNotification.toString(), slackChannelName, calendarId, calendar.getName(), event.getId(), event.getTitle(), event.getStartTime(), event.getEndTime()));
      });
    }

    if (beforeNotification && beforeNotification != '') {
      calendarEvents.forEach((event) => {
        const start = event.getStartTime();
        if (start < now.date || start.getDate() != now.date.getDate()) return;

        notifications.push(new Notification(configId, 'before', beforeNotification.toString(), slackChannelName, calendarId, calendar.getName(), event.getId(), event.getTitle(), event.getStartTime(), event.getEndTime()));
      });
    }
  });

  notificationSheet.clear();
  notificationSheet.appendRow(['Config ID', 'Type', 'Config', 'Slack Channel', 'Calendar ID', 'Calendar Title', 'Event ID', 'Event Title', 'Start', 'End', 'Eid', 'URL']);

  notifications.forEach((notification) => {
    const row = [
      notification.configId,
      notification.type,
      `="${notification.config}"`,
      notification.slackChannelName,
      notification.calendarId,
      notification.calendarTitle,
      notification.eventId,
      notification.eventTitle,
      notification.start,
      notification.end,
      notification.eid,
      notification.url,
    ];
    notificationSheet.appendRow(row);
  });
}

function notifyEvents() {
  const spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
  const notificationSheet = spreadsheet.getSheetByName('Notifications');
  const currentTick = new EventTime(new Date());
  currentTick.date.setMinutes(currentTick.date.getMinutes() - currentTick.date.getMinutes() % TICK_MINUTES);
  currentTick.date.setSeconds(0);
  currentTick.date.setMilliseconds(0);
  const nextTick = currentTick.addMinutes(TICK_MINUTES);
  const currentRange = new EventTimeRange(currentTick, nextTick);
  const isMonday = currentTick.date.getDay() == 1;

  const dailyNotifications : { [id :string] : Notification[] }= {};
  const weeklyNotifications : { [id :string] : Notification[] }= {};
  const beforeNotifications : { [id :string] : Notification[] }= {};
  notificationSheet.getRange(2, 1, notificationSheet.getLastRow() - 1, 11).getValues().forEach((row) => {
    const notification = new Notification(row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9]);

    var notificationTime;
    var notifications;
    switch (notification.type) {
      case 'daily':
        var [hour, minutes] = notification.config.split(':');
        notificationTime = currentTick.dup();
        notificationTime.date.setHours(Number(hour));
        notificationTime.date.setMinutes(Number(minutes));
        notifications = dailyNotifications;
        break
      case 'weekly':
        if (!isMonday) return;
        var [hour, minutes] = notification.config.split(':');
        notificationTime = currentTick.dup();
        notificationTime.date.setHours(Number(hour));
        notificationTime.date.setMinutes(Number(minutes));
        notifications = weeklyNotifications;
        break;
      case 'before':
        notificationTime = new EventTime(notification.start).addMinutes(-Number(notification.config));
        notifications = beforeNotifications;
        break;
      default:
        console.error(`Unknown notification type: ${notification.type}`);
        return;
    }
    if (currentRange.includes(notificationTime)) {
      if (!notifications[notification.configId]) {
        notifications[notification.configId] = [];
      }
      notifications[notification.configId].push(notification);
    }
  });

  notifySummaryNotications(dailyNotifications, 'today');
  notifySummaryNotications(weeklyNotifications, 'this week');
  notifyBeforeNotications(beforeNotifications);
}

function notifySummaryNotications(notificationsMap: { [id :string] : Notification[] }, captionRange: string) {
  const count = Object.keys(notificationsMap).reduce((sum, configId) => (sum + notificationsMap[configId].length), 0)
  if (count == 0) return;

  Object.keys(notificationsMap).forEach((configId) => {
    const notifications = notificationsMap[configId];
    const slackChannelName = notifications[0].slackChannelName;
    const calendarTitle = notifications[0].calendarTitle;
    const blocks = [{
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `There are ${count} events ${captionRange}`,
      },
    }];
    notifications.forEach((notification) => {
      const startHhmm = Utilities.formatDate(notification.start, 'Asia/Tokyo', 'HH:mm');
      const endHhmm = Utilities.formatDate(notification.end, 'Asia/Tokyo', 'HH:mm');
      const mmdd = Utilities.formatDate(notification.start, 'Asia/Tokyo', 'MM/dd');
      blocks.push({
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: `<${notification.url}|*${notification.eventTitle}*>
${  mmdd} ${startHhmm} to ${endHhmm}`,
        },
      });
    });
    notify(slackChannelName, calendarTitle, blocks);
  });
}

function notifyBeforeNotications(notifications: { [id :string] : Notification[] }) {
  Object.keys(notifications).forEach((configId) => {
    const calendarNotifications = notifications[configId];
    calendarNotifications.forEach((notification) => {
      const startHhmm = Utilities.formatDate(notification.start, 'Asia/Tokyo', 'HH:mm');
      const endHhmm = Utilities.formatDate(notification.end, 'Asia/Tokyo', 'HH:mm');
      const mmdd = Utilities.formatDate(notification.start, 'Asia/Tokyo', 'MM/dd');
      const blocks = [{
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: `Event starting at *${startHhmm}*
<${notification.url}|*${notification.eventTitle}*>
${mmdd} ${startHhmm} to ${endHhmm}`,
        },
      }];
      notify(notification.slackChannelName, notification.calendarTitle, blocks);
    });
  });
}

function notify(channel: string, username: string, blocks: any[]) {
  console.log(`Post to #${channel} from ${username}`);
  const payload = {
      token: SLACK_TOKEN,
      channel: channel,
      icon_emoji: ':calendar:',
      username: username,
      blocks: JSON.stringify(blocks),
      unfurl_links: false,
  };
  const response = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    payload,
  });
  try {
    const responseJson = JSON.parse(response.getContentText());
    if (!responseJson.ok) {
      console.error(response.getContentText());
    }
  } catch (e) {
      console.error(response.getContentText());
  }
}
