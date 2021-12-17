///
/// IMPORTS
///

import Axios = require('axios');
import { each } from 'async';
import Moment = require('moment-timezone');
import {
  Appointment,
  AppointmentAttendance,
  AppointmentAttendee,
  AppointmentNotification,
  AppointmentNotificationMethods,
  Calendar,
  ExternalCalendarPermissions,
  logger
} from 'idea-toolbox';

import { CalendarsConnector } from './calendars.connector';

///
/// CONSTANTS, ENVIRONMENT VARIABLES, HANDLER
///

const DAWN_OF_TIME = process.env['DAWN_OF_TIME'];
const END_OF_TIME = process.env['END_OF_TIME'];

const MICROSOFT_CLIENT_ID = process.env['MICROSOFT_CLIENT_ID'];
const MICROSOFT_CLIENT_SECRET = process.env['MICROSOFT_CLIENT_SECRET'];
const MICROSOFT_API_SCOPE = process.env['MICROSOFT_API_SCOPE'];
const MICROSOFT_REDIRECT_URI = process.env['MICROSOFT_REDIRECT_URI'];
const MICROSOFT_SYNC_MAX_RESULTS = Number(process.env['MICROSOFT_SYNC_MAX_RESULTS']);

const DEFAULT_CALENDAR_COLOR = '#333';

const BASE_URL_MICROSOFT_API = 'https://graph.microsoft.com/v1.0/';

///
/// CONNECTOR
///

export class MicrosoftCalendarsConnector extends CalendarsConnector {
  getAccessToken(calendar: Calendar, force?: boolean): Promise<string> {
    return new Promise((resolve, reject) => {
      if (this.token && !force) return resolve(this.token);
      // get the refresh token
      this.dynamoDB
        .get({ TableName: this.TABLES.calendarsTokens, Key: { calendarId: calendar.calendarId } })
        .then((calToken: any) => {
          // request an access token, to make API requests
          Axios.default
            .post(
              'https://login.microsoftonline.com/common/oauth2/v2.0/token',
              `client_id=${MICROSOFT_CLIENT_ID}&client_secret=${MICROSOFT_CLIENT_SECRET}&scope=${MICROSOFT_API_SCOPE}` +
                `&refresh_token=${String(calToken.token)}&grant_type=refresh_token`,
              { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
            )
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
      const redirectURI = projectURL.concat('/', MICROSOFT_REDIRECT_URI);
      Axios.default
        .post(
          'https://login.microsoftonline.com/common/oauth2/v2.0/token',
          `client_id=${MICROSOFT_CLIENT_ID}&client_secret=${MICROSOFT_CLIENT_SECRET}&scope=${MICROSOFT_API_SCOPE}` +
            `&code=${code}&grant_type=authorization_code&redirect_uri=${redirectURI}`,
          { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
        )
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
          const url = BASE_URL_MICROSOFT_API.concat(`me/calendars/${calendar.external.calendarId}`);
          Axios.default
            .get(url, { headers: { Authorization: token } })
            .then((res: Axios.AxiosResponse) => {
              const extCal: any = res.data;
              // update the resource with the external calendar configuration
              if (!calendar.name || calendar.name === '-') calendar.name = extCal.name;
              calendar.external.name = extCal.name;
              if (extCal.canShare) calendar.external.userAccess = ExternalCalendarPermissions.OWNER;
              else if (extCal.canEdit) calendar.external.userAccess = ExternalCalendarPermissions.WRITER;
              else if (extCal.canViewPrivateItems) calendar.external.userAccess = ExternalCalendarPermissions.READER;
              else calendar.external.userAccess = ExternalCalendarPermissions.FREE_BUSY;
              if (!calendar.color) calendar.color = extCal.color === 'auto' ? DEFAULT_CALENDAR_COLOR : extCal.color;
              // get the email the user used to register to the service
              const urlProfile = BASE_URL_MICROSOFT_API.concat('me');
              Axios.default
                .get(urlProfile, { headers: { Authorization: 'Bearer '.concat(token) } })
                .then((p: Axios.AxiosResponse) => {
                  const profile = p.data;
                  // add the service email to the external info; useful for managing event invitations
                  calendar.external.email = profile.mail || profile.userPrincipalName;
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
    return new Promise((resolve, reject) => {
      logger('SYNC CALENDAR', null, firstSync ? 'FIRST SYNC' : 'DELTA');
      this.getAccessToken(calendar)
        .then(token => {
          // start the sync from last bookmark (delta); if there isn't one, start frorm fresh with a large date interval
          const url =
            calendar.external.syncBookmark ||
            BASE_URL_MICROSOFT_API.concat(
              `me/calendars/${calendar.external.calendarId}/calendarView/delta`,
              `?startDateTime=${DAWN_OF_TIME}`,
              `&endDateTime=${END_OF_TIME}`
            );
          // run the request
          Axios.default
            .get(url, { headers: { Authorization: token, Prefer: `odata.maxpagesize=${MICROSOFT_SYNC_MAX_RESULTS}` } })
            .then((res: Axios.AxiosResponse) => {
              // if there was data to sync, update the sync bookmark
              // -> nextLink: there is more data; deltaLink: there was data but after this run it's synchronised
              const newBookmark = res.data['@odata.nextLink'] || res.data['@odata.deltaLink'];
              logger('IS THERE MORE DATA AFTER THIS RUN?', null, Boolean(res.data['@odata.nextLink']));
              // gather the data to sync, if any
              const values = res.data.value;
              logger('APPOINTMENTS TO SYNC', null, values.length);
              if (!values.length) return resolve(false); // no appointments to update
              // prepare the support structures
              const appToRemove = new Array<string>();
              const occurrences: { [key: string]: any[] } = {}; // grouped by master appointment
              const appointments = new Array<Appointment>();
              // manage each value based on its type
              values.forEach((x: any) => {
                // in case the appointment was removed, flag it for later
                if (x['@removed']) appToRemove.push(x.id);
                // in case the appointment is an occurrence (recurrent), flag it for later
                else if (x.type === 'occurrence') {
                  if (occurrences[x.seriesMasterId]) occurrences[x.seriesMasterId].push(x);
                  else occurrences[x.seriesMasterId] = [x];
                } else {
                  // "normal" appointment or exceptions to occurences (note: exceptions are considered "full events")
                  appointments.push(this.convertAppointmentFromExternal(x, calendar));
                }
              });
              // since the edited/deleted occurences aren't managed as normal events by Microsoft,
              // we need to delete/re-insert all of them every time one of them changes.
              // Note: when an occurence changes, all of its siblings are returned in the same API request.
              // This is a parallel action (for each master event) that waits for its occurrences to be (re)inserted.
              const masterAppointmentsIds = Object.keys(occurrences);
              logger('MASTER APPOINTMENTS TO MANAGE', null, masterAppointmentsIds.length);
              each(
                masterAppointmentsIds,
                (masterId, done) => {
                  logger('--- OCCURENCES', null, `${occurrences[masterId].length} (${masterId})`);
                  // find the master to gather the full information
                  const masterEvent = values.find((master: any) => master.id === masterId);
                  if (!masterEvent) return done();
                  // frequent occurences are limited to avoid crashes
                  let limit = 10 * 365 * 24 * 60 * 60 * 1000; // default: ~ten years
                  switch (masterEvent.recurrence.pattern.type) {
                    case 'weekly':
                      limit = 365 * 24 * 60 * 60 * 1000; // default: ~one year
                      break;
                    case 'daily':
                      limit = 60 * 24 * 60 * 60 * 1000; // default: ~two months
                      break;
                  }
                  // sort the occurrences by startDate and filter to limit the number of occurrences
                  const today = Date.now(),
                    lowerLimit = today - limit,
                    upperLimit = today + limit;
                  // prepare the occurences to insert for this master event
                  const occurrencesToInsert = occurrences[masterId]
                    .filter(
                      x =>
                        new Date(x.start.dateTime).getTime() > lowerLimit &&
                        new Date(x.start.dateTime).getTime() < upperLimit
                    )
                    .map((x: any) => this.convertAppointmentFromExternal(x, calendar, masterEvent));
                  // find and delete the previous occurences, before to insert the new ones (no error check)
                  this.deleteOccurrencesOfMasterAppointment(calendar, masterId).then(
                    () =>
                      this.batchGetPutHelper(occurrencesToInsert, firstSync)
                        .then(() => done())
                        .catch(() => done()) // ignore errors
                  );
                },
                () => {
                  // batch-save the normal appointments
                  logger('APPOINTMENTS TO INSERT', null, appointments.length);
                  this.batchGetPutHelper(appointments, firstSync).then(() => {
                    // remove the deleted appointments (flagged earlier)
                    logger('APPOINTMENTS TO DELETE', null, appToRemove.length);
                    this.dynamoDB
                      .batchDelete(
                        this.TABLES.appointments,
                        appToRemove.map(appointmentId => ({
                          calendarId: calendar.getCalendarIdForAppointments(),
                          appointmentId
                        }))
                      )
                      .then(() =>
                        // delete the occurrences of the master events deleted (if any)
                        each(
                          appToRemove,
                          (maybeMasterId, done) =>
                            this.deleteOccurrencesOfMasterAppointment(calendar, maybeMasterId).then(() => done()),
                          () => {
                            // update the calendar to bookmark the current sync status
                            calendar.external.syncBookmark = newBookmark;
                            calendar.external.lastSyncAt = Date.now();
                            this.dynamoDB
                              .put({ TableName: this.TABLES.calendars, Item: calendar })
                              // return true if finished, false if there is still data to sync
                              .then(() => resolve(Boolean(res.data['@odata.nextLink'])))
                              .catch((err: Error) => reject(err));
                          }
                        )
                      )
                      .catch((err: Error) => reject(err));
                  });
                }
              );
            })
            .catch((err: Error) => reject(err));
        })
        .catch((err: Error) => reject(err));
    });
  }
  /**
   * Delete all the occurences of a master appointment (-> it has a recurrence).
   */
  protected deleteOccurrencesOfMasterAppointment(calendar: Calendar, masterId: string): Promise<void> {
    return new Promise(resolve => {
      // find all the previous occurences
      this.dynamoDB
        .query({
          TableName: this.TABLES.appointments,
          IndexName: 'calendarId-masterAppointmentId-index',
          KeyConditionExpression: 'calendarId = :calendarId AND masterAppointmentId = :masterAppointmentId',
          ExpressionAttributeValues: {
            ':calendarId': calendar.getCalendarIdForAppointments(),
            ':masterAppointmentId': masterId
          }
        })
        .then(
          (occurrencesToDelete: Appointment[]) =>
            // delete the occurences
            this.dynamoDB
              .batchDelete(
                this.TABLES.appointments,
                occurrencesToDelete.map(a => ({
                  calendarId: calendar.getCalendarIdForAppointments(),
                  appointmentId: a.appointmentId
                }))
              )
              .then(() => resolve())
              .catch(() => resolve()) // ignore error
        )
        .catch(() => resolve()); // ignore error
    });
  }

  getAppointment(calendar: Calendar, appointmentId: string): Promise<Appointment> {
    return new Promise((resolve, reject) => {
      this.getAccessToken(calendar)
        .then(token => {
          // get and return the external appointment
          const url = BASE_URL_MICROSOFT_API.concat(
            `me/calendars/${calendar.external.calendarId}/events/${appointmentId}`
          );
          Axios.default
            .get(url, { headers: { Authorization: token } })
            .then((res: Axios.AxiosResponse) => resolve(this.convertAppointmentFromExternal(res.data, calendar)))
            .catch((err: Error) => reject(err));
        })
        .catch((err: Error) => reject(err));
    });
  }

  postAppointment(calendar: Calendar, appointment: Appointment): Promise<Appointment> {
    return new Promise((resolve, reject) => {
      this.getAccessToken(calendar)
        .then(token => {
          // prepare the appointment in Microsoft's format
          const app = this.convertAppointmentToExternal(appointment);
          // request the creation of the new appointment
          const url = BASE_URL_MICROSOFT_API.concat(`me/calendars/${calendar.external.calendarId}/events`);
          Axios.default
            .post(url, app, { headers: { Authorization: token } })
            .then((res: Axios.AxiosResponse) => resolve(this.convertAppointmentFromExternal(res.data, calendar)))
            .catch((err: Error) => {
              logger('POST APPOINTMENT', err);
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
          // prepare the appointment in Microsoft's format
          const app = this.convertAppointmentToExternal(appointment);
          // request the edit of the appointment
          const url = BASE_URL_MICROSOFT_API.concat(`me/events/${appointment.appointmentId}`);
          Axios.default
            .patch(url, app, { headers: { Authorization: token } })
            .then(() => resolve())
            .catch((err: Error) => {
              logger('PUT APPOINTMENT', err);
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
          const url = BASE_URL_MICROSOFT_API.concat(`me/events/${appointmentId}`);
          Axios.default
            .delete(url, { headers: { Authorization: token } })
            .then(() => resolve())
            .catch((err: Error) => {
              logger('DELETE APPOINTMENT', err);
              reject(err);
            });
        })
        .catch((err: Error) => reject(err));
    });
  }

  /**
   * Note: Microsoft doesn't allow changing the attendees statuses with the classic patch;
   * it requires a specific request.
   */
  updateAppointmentAttendance(
    calendar: Calendar,
    appointment: Appointment,
    attendance: AppointmentAttendance
  ): Promise<void> {
    return new Promise((resolve, reject) => {
      this.getAccessToken(calendar)
        .then(token => {
          // calculate the request URL based on the attendance
          const reqURL = this.getRequestURLByAttendance(attendance, appointment.appointmentId);
          // request the edit of the appointment
          Axios.default
            .post(reqURL, { sendResponse: true }, { headers: { Authorization: 'Bearer '.concat(token) } })
            .then(() => resolve())
            .catch((err: Error) => {
              logger('PATCH APPOINTMENT', err);
              reject(err);
            });
        })
        .catch((err: Error) => reject(err));
    });
  }

  ///
  /// DATA MAPPING
  ///

  private convertAppointmentFromExternal(x: any, calendar: Calendar, masterEvent?: any): Appointment {
    return new Appointment({
      appointmentId: x.id,
      calendarId: calendar.getCalendarIdForAppointments(),
      iCalUID: x.iCalUId,
      masterAppointmentId: masterEvent ? masterEvent.id : undefined,
      title: masterEvent ? masterEvent.subject : x.subject,
      location: masterEvent ? masterEvent.location.displayName : x.location.displayName,
      description: masterEvent ? masterEvent.bodyPreview : x.bodyPreview,
      // note: allDay appointments are set midnight to midnight (of the day after)
      startTime: x.isAllDay
        ? Number(Moment.tz(x.start.dateTime, x.start.timeZone).startOf('day').format('x'))
        : Number(Moment.tz(x.start.dateTime, x.start.timeZone).format('x')),
      // note: allDay appointments are set midnight to midnight (of the day after)
      endTime: x.isAllDay
        ? Number(Moment.tz(x.end.dateTime, x.end.timeZone).startOf('day').subtract(11, 'hours').format('x'))
        : Number(Moment.tz(x.end.dateTime, x.end.timeZone).format('x')),
      allDay: masterEvent ? masterEvent.isAllDay : x.isAllDay,
      timezone: x.start.timeZone,
      linkToOrigin: x.webLink,
      notifications: x.isReminderOn
        ? [
            new AppointmentNotification({
              method: AppointmentNotificationMethods.PUSH,
              minutes: x.reminderMinutesBeforeStart
            })
          ]
        : [],
      attendees: ((masterEvent ? masterEvent.attendees : x.attendees) || []).map(
        (a: any) =>
          new AppointmentAttendee({
            email: a.emailAddress.address,
            organizer:
              a.emailAddress.address === (masterEvent ? masterEvent.organizer.emailAddress : x.organizer.emailAddress),
            self: a.emailAddress.address === calendar.external.email,
            attendance: this.convertAttendanceFromExternal(a.status.response)
          })
      ),
      calendar
    });
  }
  protected convertAppointmentToExternal(appointment: Appointment): any {
    const res: any = {
      subject: appointment.title,
      location: { displayName: appointment.location },
      body: { content: appointment.description, contentType: 'text' },
      // note: allDay appointments are set midnight to midnight (of the day after)
      start: {
        dateTime: appointment.allDay
          ? Moment(appointment.startTime).startOf('day').format()
          : Moment(appointment.startTime).format(),
        timeZone: appointment.timezone
      },
      end: {
        dateTime: appointment.allDay
          ? Moment(appointment.endTime).add(1, 'day').startOf('day').format()
          : Moment(appointment.endTime).format(),
        timeZone: appointment.timezone
      },
      isAllDay: appointment.allDay,
      isReminderOn: Boolean(appointment.notifications.length),
      attendees: appointment.attendees.map(a => ({
        status: { response: this.convertAttendanceToExternal(a.attendance) },
        emailAddress: { address: a.email }
      }))
    };
    // needed since Microsoft does not accept reminderMinutesBeforeStart: null
    if (appointment.notifications.length) res.reminderMinutesBeforeStart = appointment.notifications[0].minutes;
    return res;
  }
  private getRequestURLByAttendance(attendance: AppointmentAttendance, id: string): string {
    const url = BASE_URL_MICROSOFT_API.concat(`me/events/${id}/`);
    switch (attendance) {
      case AppointmentAttendance.DECLINED:
        return url.concat('decline');
      case AppointmentAttendance.TENTATIVE:
        return url.concat('tentativelyAccept');
      case AppointmentAttendance.ACCEPTED:
        return url.concat('accept');
    }
  }
  private convertAttendanceToExternal(attendance: AppointmentAttendance): string {
    switch (attendance) {
      case AppointmentAttendance.NEEDS_ACTION:
        return 'none';
      case AppointmentAttendance.DECLINED:
        return 'declined';
      case AppointmentAttendance.TENTATIVE:
        return 'tentativelyAccepted';
      case AppointmentAttendance.ACCEPTED:
        return 'accepted';
    }
  }
  private convertAttendanceFromExternal(external: string): AppointmentAttendance {
    switch (external) {
      case 'none':
        return AppointmentAttendance.NEEDS_ACTION;
      case 'declined':
        return AppointmentAttendance.DECLINED;
      case 'tentativelyAccepted':
        return AppointmentAttendance.TENTATIVE;
      case 'accepted':
        return AppointmentAttendance.ACCEPTED;
    }
  }
}
