/* eslint-disable react/self-closing-comp */
import * as React from "react";
import { ISpAggregatedCalendarProps } from "./ISpAggregatedCalendarProps";
import styles from "./SpAggregatedCalendar.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";
import * as moment from "moment";
import FullCalendar from "@fullcalendar/react";
import rrulePlugin from "@fullcalendar/rrule";
import timeGridPlugin from "@fullcalendar/timegrid";
import dayGridPlugin from "@fullcalendar/daygrid";
import listPlugin from "@fullcalendar/list";
import {
  Calendar,
  CommandBar,
  defaultCalendarStrings,
  getTheme,
  IStackTokens,
  Stack,
  Spinner,
  SpinnerSize
} from "@fluentui/react";
import * as strings from "SpAggregatedCalendarWebPartStrings";
import { SpCalendarLegend } from "../calendarLegend/SpCalendarLegend";
import { IEventData } from "../../../services/IEventData";
import spservices from "../../../services/spservices";
import { timeToUntilString } from "rrule/dist/esm/dateutil";


const containerStackTokens: IStackTokens = { childrenGap: 5 };
const verticalGapStackTokens: IStackTokens = {
  childrenGap: 10,
  padding: 10,
};
const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 5,
  padding: 10,
};
const clickableStackTokens: IStackTokens = {
  padding: 10,
};

export enum formTypes {
  new = 1,
  edit = 2,
  delete = 3,
}

export const SpAggregatedCalendar: React.FunctionComponent<
  ISpAggregatedCalendarProps
> = (props: ISpAggregatedCalendarProps) => {
  const [isCalloutVisible, setCalloutVisible] = React.useState(false);
  const [selectedCalendarList, setSelectedCalendarList] = React.useState([]);
  const [calendarEvents, setCalendarEvents] = React.useState([]);
  const [selectedEvent, setSelectedEvent] = React.useState({
    siteURL: "",
    calendarName: "",
    id: "0",
    title: "",
    backgroundColor: "",
    start: moment().toISOString(),
    end: moment().toISOString(),
    description: "",
    location: "",
    allDay: false,
    category: "",
    RecurrenceData: "",
    fRecurrence: false,
  });
  const [menuItems, setMenuItems] = React.useState([]);
  const [filteredLists, setFilteredList] = React.useState([]);
  const [isModalOpen, setIsModalOpen] = React.useState(false);
  const [isDialogHidden, setIsDialogHidden] = React.useState(true);
  const [selectedListId, setSelectedListId] = React.useState("");
  const [formTypeControl, setFormTypeControl] = React.useState(formTypes.new);
  const [menuListItems, setMenuListItems] = React.useState([]);
  const [selectedListTitlle, setSelectedListTitle] = React.useState("");
  const [isLoading, calendarIsLoading]=React.useState(false);
  const [eventSourcesArray, setEvents] = React.useState([]);
  const [viewDateRange, setViewDateRange] = React.useState<{
    start: Date;
    end: Date;
  }>({
    start: moment().startOf("month").toDate(),
    end: moment().endOf("month").toDate(),
  });
  const _dataService: spservices = new spservices(props.context);
  React.useEffect(() => {
    function fetchEvents() {
      let myEvents: any = [];
      const promises = props.selectedCalendarLists.map((calendarData) =>
        _dataService
          .getEvents(
            escape(calendarData.SiteUrl),
            escape(calendarData.CalendarListTitle),
            viewDateRange.start,
            viewDateRange.end
          )
          .then((eventsData) => {
            myEvents.push({
              id: calendarData.CalendarTitle,
              events: eventsData,
              color: calendarData.Color
            });
          })
          .catch((error) => {
            console.error(error);
          })
      );

      return Promise.all(promises).then(() => myEvents); // Wait for all promises to resolve
    }

    fetchEvents()
      .then((myEvents) => {
        console.log(myEvents);
        setEvents(myEvents);
        calendarIsLoading(false);
      })
      .catch((error) => {
        console.error(error);
      });
  }, [viewDateRange, isLoading]);

  const handleDateSet = (info: any): void => {
    const { view, start, end } = info;
    console.log(`Visible range in ${view.type} view: ${start} - ${end}`);
    setViewDateRange({ start: start, end: end });
    calendarIsLoading(true);
  };

  const theme = getTheme();
  const calendarComponentRef: any = React.createRef();

  const navigateCalendar = (date: Date): void => {
    const calendarApi = calendarComponentRef.current.getApi();
    console.log(calendarApi.getEventSources());
    calendarApi.gotoDate(date);
  };

  return (
    <div className={styles.spAggregatedCalendar}>
      <Stack tokens={containerStackTokens}>
        <hr color={theme.palette.themePrimary} />
        <Stack tokens={verticalGapStackTokens}>
          <Stack horizontalAlign="end">
            <CommandBar
              items={[
                {
                  key: "newItem",
                  text: "New Event",
                  cacheKey: "newItemCache", // changing this key will invalidate this item's cache
                  iconProps: { iconName: "Add" },
                  subMenuProps: {
                    items: menuItems,
                  },
                },
                {
                  key: "exportItems",
                  text: "Export Events",
                  cacheKey: "exportItemCache", // changing this key will invalidate this item's cache
                  iconProps: { iconName: "ExcelDocument" },
                  subMenuProps: {
                    items: menuItems,
                  },
                },
              ]}
              // overflowButtonProps={overflowProps}
              ariaLabel="Calendar Commands"
              primaryGroupAriaLabel="Email actions"
              farItemsGroupAriaLabel="More actions"
            />
          </Stack>

          <Stack horizontal horizontalAlign="space-between">
            <Stack>
              <Calendar
                showMonthPickerAsOverlay
                highlightSelectedMonth
                showGoToToday={true}
                onSelectDate={(date) => navigateCalendar(date)}
                // value={selectedDate}
                // Calendar uses English strings by default. For localized apps, you must override this prop.
                strings={defaultCalendarStrings}
              />
              <Stack horizontal horizontalAlign="start">
                {" "}
                {props.showLegend && (
                  <div>
                    <p>Calendars in View:</p>
                    <div className={styles.legend}>
                      <SpCalendarLegend
                        selectedCalendarLists={props.selectedCalendarLists}
                      />
                    </div>
                  </div>
                )}
              </Stack>
            </Stack>
            <div>
              {isLoading ? <Spinner size={SpinnerSize.large} label={strings.LoadingEventsLabel}/> : <FullCalendar
                plugins={[
                  timeGridPlugin,
                  dayGridPlugin,
                  rrulePlugin,
                  listPlugin,
                ]}
                initialView={props.defaultView}
                headerToolbar={{
                  center: "title",
                  start: "timeGridDay,timeGridWeek,dayGridMonth,listMonth",
                  end: "prev,next,today",
                }}
                businessHours={{
                  // days of week. an array of zero-based day of week integers (0=Sunday)
                  daysOfWeek: [1, 2, 3, 4, 5], // Monday - Thursday

                  startTime: "08:00", // a start time (10am in this example)
                  endTime: "17:00", // an end time (6pm in this example)
                }}
                initialDate={new Date()}
                navLinks={true}
                editable={true}
                aspectRatio={2}
                ref={calendarComponentRef}
                datesSet={handleDateSet}
                // eventLimit = {3}
                fixedWeekCount={false}
                // eventClick={this.eventClickHandler}
                eventSources={eventSourcesArray}
              ></FullCalendar>}
              
            </div>
          </Stack>
        </Stack>
      </Stack>
    </div>
  );
};
