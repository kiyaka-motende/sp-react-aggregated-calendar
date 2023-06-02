/* eslint-disable no-useless-escape */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp, Web, PermissionKind, RegionalSettings } from "@pnp/sp";
import { IEventData } from "./IEventData";
import * as moment from "moment";
import { RRule } from "rrule";
import parseRecurrentEvent from "./parseRecurrentEvent";
// import { graph} from "@pnp/graph";

// Class Services
export default class spservices {
  constructor(private context: WebPartContext) {
    // Setuo Context to PnPjs and MSGraph
    sp.setup({
      spfxContext: this.context,
    });

    // graph.setup({
    //   spfxContext: this.context
    // });
    // Init
    // this.onInit();
  }

  // OnInit Function
  private async onInit() {}

  /**
   *
   * @private
   * @returns {Promise<string>}
   * @memberof spservices
   */
  public async getLocalTime(date: string | Date): Promise<string> {
    try {
      const localTime = await sp.web.regionalSettings.timeZone.utcToLocalTime(
        date
      );
      return localTime;
    } catch (error) {
      return Promise.reject(error);
    }
  }

  /**
   *
   * @private
   * @returns {Promise<string>}
   * @memberof spservices
   */
  public async getUtcTime(date: string | Date): Promise<string> {
    try {
      const utcTime = await sp.web.regionalSettings.timeZone.localTimeToUTC(
        date
      );
      return utcTime;
    } catch (error) {
      return Promise.reject(error);
    }
  }

  /**
   *
   * @param {string} siteUrl
   * @param {string} listId
   * @param {string} fieldInternalName
   * @returns {Promise<{ key: string, text: string }[]>}
   * @memberof spservices
   */
  public async getChoiceFieldOptions(
    siteUrl: string,
    listId: string,
    fieldInternalName: string
  ): Promise<{ key: string; text: string }[]> {
    let fieldOptions: { key: string; text: string }[] = [];
    try {
      const web = new Web(siteUrl);
      const results = await web.lists
        .getById(listId)
        .fields.getByInternalNameOrTitle(fieldInternalName)
        .select("Title", "InternalName", "Choices")
        .get();
      if (results && results.Choices.length > 0) {
        for (const option of results.Choices) {
          fieldOptions.push({
            key: option,
            text: option,
          });
        }
      }
    } catch (error) {
      return Promise.reject(error);
    }
    return fieldOptions;
  }

  /**
   *
   * @param {string} siteUrl
   * @param {string} listId
   * @param {Date} eventStartDate
   * @param {Date} eventEndDate
   * @returns {Promise< IEventData[]>}
   * @memberof spservices
   */
  public async getEvents(
    siteUrl: string,
    listTitle: string,
    eventStartDate: Date, 
    eventEndDate: Date
  ): Promise<IEventData[]> {
    let events: IEventData[] = [];
    if (!siteUrl) {
      return [];
    }
    const parseEvt: parseRecurrentEvent = new parseRecurrentEvent();
      
    try {
      console.log(moment(eventStartDate).format('YYYY-MM-DD'));
      const web = new Web(siteUrl);
      const results = await web.lists
        .getByTitle(listTitle)
        .usingCaching()
        .renderListDataAsStream({
          DatesInUtc: true,
          ViewXml: `<View><ViewFields><FieldRef Name='RecurrenceData'/><FieldRef Name='Duration'/><FieldRef Name='Author'/><FieldRef Name='Category'/><FieldRef Name='Description'/><FieldRef Name='ParticipantsPicker'/><FieldRef Name='Geolocation'/><FieldRef Name='ID'/><FieldRef Name='EndDate'/><FieldRef Name='EventDate'/><FieldRef Name='ID'/><FieldRef Name='Location'/><FieldRef Name='Title'/><FieldRef Name='fAllDayEvent'/><FieldRef Name='EventType'/><FieldRef Name='UID' /><FieldRef Name='fRecurrence' /></ViewFields>
          <Query>
          <Where>
            <And>
              <Geq>
                <FieldRef Name='EventDate' />
                <Value IncludeTimeValue='false' Type='DateTime'>${moment(eventStartDate).format('YYYY-MM-DD')}</Value>
              </Geq>
              <Leq>
                <FieldRef Name='EventDate' />
                <Value IncludeTimeValue='false' Type='DateTime'>${moment(eventEndDate).format('YYYY-MM-DD')}</Value>
              </Leq>
              </And>
          </Where>
          </Query>      
          <RowLimit Paged="FALSE">2000</RowLimit>
          </View>`,
        });

      if (results && results.Row.length > 0) {
        let event: any = "";
        const mapEvents = async (): Promise<boolean> => {
          for (event of results.Row) {
            const eventDate = await this.getLocalTime(event.EventDate);
            const endDate = await this.getLocalTime(event.EndDate);
            const initialsArray: string[] = event.Author[0].title.split(" ");
            const initials: string =
              initialsArray[0].charAt(0) +
              initialsArray[initialsArray.length - 1].charAt(0);
            //   const userPictureUrl = await this.getUserProfilePictureUrl(`i:0#.f|membership|${event.Author[0].email}`);
            const attendees: number[] = [];
            const first: number = event.Geolocation.indexOf("(") + 1;
            const last: number = event.Geolocation.indexOf(")");
            const geo = event.Geolocation.substring(first, last);
            const geolocation = geo.split(" ");
            //   const CategoryColorValue: any[] = categoryColor.filter((value) => {
            //     return value.category == event.Category;
            //   });
            const isAllDayEvent: boolean = event["fAllDayEvent.value"] === "1";
            const isRecurring: boolean = event["fRecurrence.value"] === "1";

            for (const attendee of event.ParticipantsPicker) {
              attendees.push(parseInt(attendee.id));
            }

            events.push({
              Id: event.ID,
              ID: event.ID,
              EventType: event.EventType,
              title: await parseEvt.deCodeHtmlEntities(event.Title),
              Description: event.Description,
              start: isAllDayEvent
                ? new Date(event.EventDate.slice(0, -1))
                : new Date(eventDate),
              end: isAllDayEvent
                ? new Date(event.EndDate.slice(0, -1))
                : new Date(endDate),
              location: event.Location,
              ownerEmail: event.Author[0].email,
              // ownerPhoto: userPictureUrl ?
              //   `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${event.Author[0].email}&UA=0&size=HR96x96` : '',
              ownerInitial: initials,
              // color: CategoryColorValue.length > 0 ? CategoryColorValue[0].color : '#1a75ff', // blue default
              ownerName: event.Author[0].title,
              attendes: attendees,
              geolocation: {
                Longitude: parseFloat(geolocation[0]),
                Latitude: parseFloat(geolocation[1]),
              },
              Category: event.Category,
              Duration: event.Duration,
              RecurrenceData:event.RecurrenceData,
              allDay:isAllDayEvent,
              rrule: isRecurring
                ? RRule.fromText(
                    event.EventType === "4" && event.MasterSeriesItemID !== ""
                      ? await parseEvt.deCodeHtmlEntities(event.RecurrenceData)
                      : await parseEvt.returnExceptionRecurrenceInfo(
                          await parseEvt.deCodeHtmlEntities(
                            event.RecurrenceData
                          )
                        )
                  ).options
                : null,
              fRecurrence: event.fRecurrence,
              RecurrenceID: event.RecurrenceID ? event.RecurrenceID : undefined,
              MasterSeriesItemID: event.MasterSeriesItemID,
              UID: event.UID.replace("{", "").replace("}", ""),
            });
          }
          return true;
        };
        await mapEvents();
      }

      //   events = parseEvt.parseEvents(events, null, null);

      // Return Data
      return events;
    } catch (error) {
      console.dir(error);
      return Promise.reject(error);
    }
  }
}
