import { each } from 'async';
import { Calendar, Appointment, AppointmentAttendance } from 'idea-toolbox';

export abstract class CalendarsConnector {
  /**
   * The access token used to perform API requests.
   */
  protected token: string;
  /**
   * The dynamoDB tables involved.
   */
  protected TABLES: any = {
    calendars: 'idea_calendars',
    appointments: 'idea_calendars_appointments',
    calendarsTokens: 'idea_externalCalendarsTokens'
  };

  /**
   * Initialise the connector with a dynamoDB instance to use and the prefix for RESPONSE_URI.
   */
  constructor(protected dynamoDB: any) {}

  /**
   * Connect the calendar with its external link.
   */
  public abstract configure(calendarId: string, code: string, projectURL: string): Promise<void>;

  /**
   * Update the configuration of the calendar with external data.
   */
  public abstract updateCalendarConfiguration(calendar: Calendar): Promise<Calendar>;

  /**
   * Get an access token for the service.
   */
  public abstract getAccessToken(calendar: Calendar, force?: boolean): Promise<string>;

  /**
   * Synchronise the given calendar with its linked one.
   */
  public abstract syncCalendar(calendar: Calendar, firstSync?: boolean): Promise<boolean>;

  /**
   * Get an appointment from the external calendar.
   */
  public abstract getAppointment(calendar: Calendar, appointmentId: string): Promise<Appointment>;

  /**
   * Add an appointment in the external calendar and return it updated with external information.
   */
  public abstract postAppointment(calendar: Calendar, appointment: Appointment): Promise<Appointment>;

  /**
   * Edit an appointment in the external calendar.
   */
  public abstract putAppointment(calendar: Calendar, appointment: Appointment): Promise<void>;

  /**
   * Delete an appointment from the external calendar.
   */
  public abstract deleteAppointment(calendar: Calendar, appointmentId: string): Promise<void>;

  /**
   * Update the attendance status of an appointment.
   */
  public abstract updateAppointmentAttendance(
    calendar: Calendar,
    appointment: Appointment,
    attendance: AppointmentAttendance
  ): Promise<void>;

  /**
   * Get/put for each appointment (one by one) to avoid erasing the linked objects; no error checking.
   */
  protected batchGetPutHelper(appointments: Array<Appointment>, directPut?: boolean): Promise<void> {
    return new Promise(resolve => {
      if (directPut) {
        this.dynamoDB
          .batchPut(this.TABLES.appointments, appointments)
          .then(() => resolve())
          .catch(() => resolve()); // ignore error
      } else {
        each(
          appointments,
          (app: Appointment, done: any) => {
            this.dynamoDB
              .get({
                TableName: this.TABLES.appointments,
                Key: { calendarId: app.calendarId, appointmentId: app.appointmentId }
              })
              .then((safeApp: Appointment) => {
                // update existing resource
                if (safeApp.linkedTo) app.linkedTo = safeApp.linkedTo;
                this.dynamoDB
                  .put({ TableName: this.TABLES.appointments, Item: app })
                  .then(() => done())
                  .catch(() => done()); // ignore error
              })
              .catch(() => {
                // save new resource
                this.dynamoDB
                  .put({ TableName: this.TABLES.appointments, Item: app })
                  .then(() => done())
                  .catch(() => done()); // ignore error
              });
          },
          () => resolve() // ignore error
        );
      }
    });
  }
}
