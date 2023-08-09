///
/// IMPORTS
///

import Axios = require('axios');
import Moment = require('moment-timezone');
import {
  Appointment,
  AppointmentAttendance,
  AppointmentAttendee,
  AppointmentNotification,
  AppointmentNotificationMethods,
  Calendar,
  ExternalCalendarPermissions
} from 'idea-toolbox';
import { CalendarsConnector } from './calendars.connector';

///
/// CONSTANTS, ENVIRONMENT VARIABLES, HANDLER
///

const GOOGLE_CLIENT_ID = process.env['GOOGLE_CLIENT_ID'];
const GOOGLE_CLIENT_SECRET = process.env['GOOGLE_CLIENT_SECRET'];
const GOOGLE_API_SCOPE = process.env['GOOGLE_API_SCOPE'];
const GOOGLE_REDIRECT_URI = process.env['GOOGLE_REDIRECT_URI'];
const GOOGLE_MAX_ATTENDEES = process.env['GOOGLE_MAX_ATTENDEES'];
const GOOGLE_SYNC_MAX_RESULTS = Number(process.env['GOOGLE_SYNC_MAX_RESULTS']);

const BASE_URL_GOOGLE_API = 'https://www.googleapis.com/';

///
/// CONNECTOR
///

export class GoogleCalendarsConnector extends CalendarsConnector {
  getAccessToken(calendar: Calendar, force?: boolean): Promise<string> {
    return new Promise((resolve, reject) => {
      if (this.token && !force) return resolve(this.token);
      // get the refresh token
      this.dynamoDB
        .get({ TableName: this.TABLES.calendarsTokens, Key: { calendarId: calendar.calendarId } })
        .then((calToken: any) => {
          // request an access token, to make API requests
          const url = BASE_URL_GOOGLE_API.concat(
            'oauth2/v4/token',
            `?client_id=${GOOGLE_CLIENT_ID}`,
            `&client_secret=${GOOGLE_CLIENT_SECRET}`,
            `&scope=${GOOGLE_API_SCOPE}`,
            `&refresh_token=${String(calToken.token)}`,
            '&grant_type=refresh_token'
          );
          Axios.default
            .post(url, { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } })
            .then((res: Axios.AxiosResponse) => {
              const refreshToken: string = res.data.refresh_token;
              const accessToken: string = res.data.access_token;
              // return the access token
              this.token = accessToken;
              resolve(accessToken);
              // (async) save the new refresh_token, if any
              if (refreshToken)
                this.dynamoDB
                  .put({
                    TableName: this.TABLES.calendarsTokens,
                    Item: { calendarId: calendar.calendarId, token: refreshToken }
                  })
                  .catch(() => {
                    /* ignore errors */
                  });
            })
            .catch((err: Error) => reject(err));
        })
        .catch((err: Error) => reject(err));
    });
  }

  configure(calendarId: string, code: string, projectURL: string): Promise<void> {
    return new Promise((resolve, reject) => {
      // send the code provided to validate it and to receive the tokens
      const url = BASE_URL_GOOGLE_API.concat(
        'oauth2/v4/token',
        `?client_id=${GOOGLE_CLIENT_ID}`,
        `&client_secret=${GOOGLE_CLIENT_SECRET}`,
        `&scope=${GOOGLE_API_SCOPE}`,
        `&code=${code}`,
        '&grant_type=authorization_code',
        `&redirect_uri=${projectURL.concat('/', GOOGLE_REDIRECT_URI)}`
      );
      Axios.default
        .post(url, { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } })
        .then((res: Axios.AxiosResponse) => {
          const refreshToken: string = res.data.refresh_token;
          // get the calendar
          this.dynamoDB
            .get({ TableName: this.TABLES.calendars, Key: { calendarId } })
            .then((safeCalendar: Calendar) => {
              // stop the process if the calendar was already configured
              if (safeCalendar.external.calendarId) return reject(new Error('CALENDAR_WAS_ALREADY_CONFIGURED'));
              // save the (long-living) token for accessing remote data
              this.dynamoDB
                .put({ TableName: this.TABLES.calendarsTokens, Item: { calendarId, token: refreshToken } })
                .then(() => resolve())
                .catch((err: Error) => reject(err));
            })
            .catch((err: Error) => reject(err));
        })
        .catch((err: Error) => reject(err));
    });
  }

  updateCalendarConfiguration(calendar: Calendar): Promise<Calendar> {
    return new Promise((resolve, reject) => {
      this.getAccessToken(calendar)
        .then(token => {
          // get the external calendar details
          const url = BASE_URL_GOOGLE_API.concat(`calendar/v3/users/me/calendarList/${calendar.external.calendarId}`);
          Axios.default
            .get(url, { headers: { Authorization: 'Bearer '.concat(token) } })
            .then((res: Axios.AxiosResponse) => {
              const extCal: any = res.data;
              // update the resource with the external calendar configuration
              if (!calendar.name || calendar.name === '-') calendar.name = extCal.summary;
              calendar.timezone = extCal.timeZone;
              if (!calendar.color) calendar.color = extCal.backgroundColor;
              calendar.external.name = extCal.summary;
              switch (extCal.accessRole) {
                case 'owner':
                  calendar.external.userAccess = ExternalCalendarPermissions.OWNER;
                  break;
                case 'writer':
                  calendar.external.userAccess = ExternalCalendarPermissions.WRITER;
                  break;
                case 'reader':
                  calendar.external.userAccess = ExternalCalendarPermissions.READER;
                  break;
                default:
                  calendar.external.userAccess = ExternalCalendarPermissions.FREE_BUSY;
              }
              // get the email the user used to register to the service
              const urlProfile = BASE_URL_GOOGLE_API.concat('oauth2/v1/userinfo');
              Axios.default
                .get(urlProfile, { headers: { Authorization: 'Bearer '.concat(token) } })
                .then((p: Axios.AxiosResponse) => {
                  const profile = p.data;
                  // add the service email to the external info; useful for managing event invitations
                  calendar.external.email = profile.email;
                  // save the updated resource
                  this.dynamoDB
                    .put({ TableName: this.TABLES.calendars, Item: calendar })
                    .then(() => resolve(calendar))
                    .catch((err: Error) => reject(err));
                })
                .catch((err: Error) => reject(err));
            })
            .catch((err: Error) => reject(err));
        })
        .catch((err: Error) => reject(err));
    });
  }

  syncCalendar(calendar: Calendar, firstSync?: boolean): Promise<boolean> {
    return new Promise((resolve, reject): void => {
      console.log('SYNC CALENDAR', firstSync ? 'FIRST SYNC' : 'DELTA');
      this.getAccessToken(calendar)
        .then(token => {
          // start the sync from last bookmark; if there isn't one, start frorm fresh
          let url = BASE_URL_GOOGLE_API.concat(
            `calendar/v3/calendars/${calendar.external.calendarId}/events`,
            `?maxResults=${GOOGLE_SYNC_MAX_RESULTS}`,
            `&maxAttendees=${GOOGLE_MAX_ATTENDEES}`,
            '&showHiddenInvitations=true',
            '&singleEvents=true'
          );
          if (calendar.external.syncBookmark) url = url.concat('&syncToken=', calendar.external.syncBookmark);
          // if the pageBookmark is set it means we are synchronising another page of the same sync window
          if (calendar.external.pageBookmark) url = url.concat('&pageToken=', calendar.external.pageBookmark);
          // run the request
          Axios.default
            .get(url, { headers: { Authorization: 'Bearer '.concat(token) } })
            .then((res: Axios.AxiosResponse) => {
              // if there was data to sync, update the sync bookmark
              // -> nextPageToken: there is more data; nextSyncToken: there was data but after this run it's up to date
              const syncBookmark = res.data['nextSyncToken'] || null;
              const pageBookmark = res.data['nextPageToken'] || null;
              console.log('IS THERE MORE DATA AFTER THIS RUN?', Boolean(pageBookmark));
              // gather the data to sync, if any
              const values = res.data.items;
              console.log('APPOINTMENTS TO SYNC', values.length);
              if (!values.length) return resolve(false); // no appointments to update
              // prepare the support structures
              const appToRemove = new Array<string>();
              const appointments = new Array<Appointment>();
              // manage each value based on its type
              values.forEach((x: any) => {
                // in case the appointment was removed, flag it for later
                if (x.status === 'cancelled') appToRemove.push(x.id);
                // otherwise, transform the appointment to the current format and add it to the lsit
                else appointments.push(this.convertAppointmentFromExternal(x, calendar));
              });
              // batch-save the appointments
              console.log('APPOINTMENTS TO INSERT', appointments.length);
              this.batchGetPutHelper(appointments, firstSync)
                .then(() => {
                  // remove the deleted appointments (flagged earlier)
                  console.log('APPOINTMENTS TO DELETE', appToRemove.length);
                  this.dynamoDB
                    .batchDelete(
                      this.TABLES.appointments,
                      appToRemove.map(appointmentId => ({
                        calendarId: calendar.getCalendarIdForAppointments(),
                        appointmentId
                      }))
                    )
                    .then(() => {
                      // update the calendar to bookmark the current sync status
                      calendar.external.syncBookmark = syncBookmark;
                      calendar.external.pageBookmark = pageBookmark;
                      calendar.external.lastSyncAt = Date.now();
                      this.dynamoDB
                        .put({ TableName: this.TABLES.calendars, Item: calendar })
                        // return true if finished, false if there is still data to sync
                        .then(() => resolve(Boolean(pageBookmark)))
                        .catch((err: Error) => reject(err));
                    })
                    .catch((err: Error) => reject(err));
                })
                .catch((err: Error) => reject(err));
            })
            .catch((err: any) => {
              // Google has a specific code to force a full sync, when necessary
              if (err && err.response && err.response.status === 410) {
                // reset the syncBookmark to force a full sync at the next request
                calendar.external.syncBookmark = null;
                calendar.external.pageBookmark = null;
                calendar.external.lastSyncAt = Date.now();
                this.dynamoDB
                  .put({ TableName: this.TABLES.calendars, Item: calendar })
                  .then(() => resolve(true))
                  .catch((error: Error) => reject(error));
              } else reject(err);
            });
        })
        .catch((err: Error) => reject(err));
    });
  }

  getAppointment(calendar: Calendar, appointmentId: string): Promise<Appointment> {
    return new Promise((resolve, reject) => {
      this.getAccessToken(calendar)
        .then(token => {
          // get and return the external appointment
          const url = BASE_URL_GOOGLE_API.concat(
            `calendar/v3/calendars/${calendar.external.calendarId}/events/${appointmentId}`
          );
          Axios.default
            .get(url, { headers: { Authorization: 'Bearer '.concat(token) } })
            .then((res: Axios.AxiosResponse) => {
              if (res.data.status === 'cancelled') reject(new Error('EVENT_IS_CANCELLED'));
              else resolve(this.convertAppointmentFromExternal(res.data, calendar));
            })
            .catch((err: Error) => reject(err));
        })
        .catch((err: Error) => reject(err));
    });
  }

  postAppointment(calendar: Calendar, appointment: Appointment): Promise<Appointment> {
    return new Promise((resolve, reject) => {
      this.getAccessToken(calendar)
        .then(token => {
          // prepare the appointment in Google's format
          const app = this.convertAppointmentToExternal(appointment);
          // request the creation of the new appointment
          const url = BASE_URL_GOOGLE_API.concat(`calendar/v3/calendars/${calendar.external.calendarId}/events`);
          Axios.default
            .post(url, app, { headers: { Authorization: 'Bearer '.concat(token) } })
            .then((res: Axios.AxiosResponse) => resolve(this.convertAppointmentFromExternal(res.data, calendar)))
            .catch((err: Error) => {
              console.error('POST APPOINTMENT', err);
              reject(err);
            });
        })
        .catch((err: Error) => reject(err));
    });
  }

  putAppointment(calendar: Calendar, appointment: Appointment): Promise<void> {
    return new Promise((resolve, reject) => {
      this.getAccessToken(calendar)
        .then(token => {
          // prepare the appointment in Google's format
          const app = this.convertAppointmentToExternal(appointment);
          const id = appointment.appointmentId;
          // request the edit of the appointment
          const url = BASE_URL_GOOGLE_API.concat(
            `calendar/v3/calendars/${calendar.external.calendarId}/events/${id}`,
            '?sendUpdates=all'
          );
          Axios.default
            .patch(url, app, { headers: { Authorization: 'Bearer '.concat(token) } })
            .then(() => resolve())
            .catch((err: Error) => {
              console.error('PUT APPOINTMENT', err);
              reject(err);
            });
        })
        .catch((err: Error) => reject(err));
    });
  }

  deleteAppointment(calendar: Calendar, appointmentId: string): Promise<void> {
    return new Promise((resolve, reject) => {
      this.getAccessToken(calendar)
        .then(token => {
          // request the creation of the new appointment
          const url = BASE_URL_GOOGLE_API.concat(
            `calendar/v3/calendars/${calendar.external.calendarId}/events/${appointmentId}`
          );
          Axios.default
            .delete(url, { headers: { Authorization: 'Bearer '.concat(token) } })
            .then(() => resolve())
            .catch((err: Error) => {
              console.error('DELETE APPOINTMENT', err);
              reject(err);
            });
        })
        .catch((err: Error) => reject(err));
    });
  }

  updateAppointmentAttendance(
    calendar: Calendar,
    appointment: Appointment,
    attendance: AppointmentAttendance
  ): Promise<void> {
    // set the new status for the attendee
    const attendee = appointment.attendees.find(x => calendar.external.email === x.email);
    if (attendee) attendee.attendance = attendance;
    // update the appointment
    return this.putAppointment(calendar, appointment);
  }

  ///
  /// DATA MAPPING
  ///

  private convertAppointmentFromExternal(x: any, calendar: Calendar): Appointment {
    const app = new Appointment(
      {
        appointmentId: x.id,
        calendarId: calendar.getCalendarIdForAppointments(),
        iCalUID: x.iCalUID,
        title: x.summary || '?', // ExternalCalendarPermissions.FREE_BUSY
        location: x.location,
        description: x.description,
        startTime: x.start.date
          ? Number(Moment(x.start.date.concat('T00:00:00.000Z')).format('x'))
          : Number(Moment.tz(x.start.dateTime, x.start.timeZone || calendar.timezone).format('x')),
        endTime: x.end.date
          ? Number(Moment(x.end.date.concat('T00:00:00.000Z')).startOf('day').subtract(11, 'hours').format('x'))
          : Number(Moment.tz(x.end.dateTime, x.end.timeZone || calendar.timezone).format('x')),
        allDay: Boolean(x.start.date),
        timezone: x.start.timeZone || calendar.timezone,
        linkToOrigin: x.htmlLink,
        notifications: (x.reminders.overrides || []).map(
          (o: any) =>
            new AppointmentNotification({
              method: this.convertNotificationMethodFromExternal(o.method),
              minutes: o.minutes
            })
        ),
        attendees: (x.attendees || []).map(
          (a: any) =>
            new AppointmentAttendee({
              email: a.email,
              organizer: a.organizer,
              self: a.self,
              attendance: this.convertAttendanceFromExternal(a.responseStatus)
            })
        )
      },
      calendar
    );
    // keep a track of the masterAppointmentId, in case it's an occurrence (the repetition of an event)
    if (x.recurringEventId) app.masterAppointmentId = x.recurringEventId;
    // return the transformed appointment
    return app;
  }
  private convertAppointmentToExternal(appointment: Appointment): any {
    // note: allDay appointments need to be set midnight to midnight (of the day after)
    const start: any = { timeZone: appointment.timezone },
      end: any = { timeZone: appointment.timezone };
    if (appointment.allDay) {
      start.date = Moment(appointment.startTime).startOf('day').format('YYYY-MM-DD');
      start.dateTime = null;
      end.date = Moment(appointment.endTime).add(1, 'day').startOf('day').format('YYYY-MM-DD');
      end.dateTime = null;
    } else {
      start.dateTime = Moment(appointment.startTime).toISOString();
      start.date = null;
      end.dateTime = Moment(appointment.endTime).toISOString();
      end.date = null;
    }
    // if notifications is empty it means we are using deafult settings for Google Calendar
    const reminders = !appointment.notifications.length
      ? { useDefault: true }
      : {
          useDefault: false,
          overrides: appointment.notifications.map((n: AppointmentNotification) => ({
            method: this.convertNotificationMethodToExternal(n.method),
            minutes: n.minutes
          }))
        };
    const attendees = appointment.attendees.map(a => ({
      email: a.email,
      organizer: a.organizer,
      self: a.self,
      responseStatus: this.convertAttendanceToExternal(a.attendance)
    }));
    return {
      summary: appointment.title,
      location: appointment.location,
      description: appointment.description,
      start,
      end,
      reminders,
      attendees
    };
  }
  private convertNotificationMethodToExternal(method: AppointmentNotificationMethods): string {
    switch (method) {
      case AppointmentNotificationMethods.PUSH:
        return 'popup';
      case AppointmentNotificationMethods.EMAIL:
        return 'email';
    }
  }
  private convertNotificationMethodFromExternal(external: string): AppointmentNotificationMethods {
    switch (external) {
      case 'popup':
        return AppointmentNotificationMethods.PUSH;
      case 'email':
        return AppointmentNotificationMethods.EMAIL;
    }
  }
  private convertAttendanceToExternal(attendance: AppointmentAttendance): string {
    switch (attendance) {
      case AppointmentAttendance.NEEDS_ACTION:
        return 'needsAction';
      case AppointmentAttendance.DECLINED:
        return 'declined';
      case AppointmentAttendance.TENTATIVE:
        return 'tentative';
      case AppointmentAttendance.ACCEPTED:
        return 'accepted';
    }
  }
  private convertAttendanceFromExternal(external: string): AppointmentAttendance {
    switch (external) {
      case 'needsAction':
        return AppointmentAttendance.NEEDS_ACTION;
      case 'declined':
        return AppointmentAttendance.DECLINED;
      case 'tentative':
        return AppointmentAttendance.TENTATIVE;
      case 'accepted':
        return AppointmentAttendance.ACCEPTED;
    }
  }
}
