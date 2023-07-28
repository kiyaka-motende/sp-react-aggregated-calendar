/* eslint-disable react/self-closing-comp */
import * as React from "react";
import styles from "./SpAggregatedCalendar.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";
import * as moment from "moment";
import FullCalendar from "@fullcalendar/react";
import rrulePlugin from "@fullcalendar/rrule";
import timeGridPlugin from "@fullcalendar/timegrid";
import dayGridPlugin from "@fullcalendar/daygrid";
import listPlugin from "@fullcalendar/list";
import { Calendar, CommandBar, defaultCalendarStrings, getTheme, Stack, Spinner, SpinnerSize } from "@fluentui/react";
import * as strings from "SpAggregatedCalendarWebPartStrings";
import { SpCalendarLegend } from "../calendarLegend/SpCalendarLegend";
import spservices from "../../../services/spservices";
var containerStackTokens = { childrenGap: 5 };
var verticalGapStackTokens = {
    childrenGap: 10,
    padding: 10,
};
var itemAlignmentsStackTokens = {
    childrenGap: 5,
    padding: 10,
};
var clickableStackTokens = {
    padding: 10,
};
export var formTypes;
(function (formTypes) {
    formTypes[formTypes["new"] = 1] = "new";
    formTypes[formTypes["edit"] = 2] = "edit";
    formTypes[formTypes["delete"] = 3] = "delete";
})(formTypes || (formTypes = {}));
export var SpAggregatedCalendar = function (props) {
    var _a = React.useState(false), isCalloutVisible = _a[0], setCalloutVisible = _a[1];
    var _b = React.useState([]), selectedCalendarList = _b[0], setSelectedCalendarList = _b[1];
    var _c = React.useState([]), calendarEvents = _c[0], setCalendarEvents = _c[1];
    var _d = React.useState({
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
    }), selectedEvent = _d[0], setSelectedEvent = _d[1];
    var _e = React.useState([]), menuItems = _e[0], setMenuItems = _e[1];
    var _f = React.useState([]), filteredLists = _f[0], setFilteredList = _f[1];
    var _g = React.useState(false), isModalOpen = _g[0], setIsModalOpen = _g[1];
    var _h = React.useState(true), isDialogHidden = _h[0], setIsDialogHidden = _h[1];
    var _j = React.useState(""), selectedListId = _j[0], setSelectedListId = _j[1];
    var _k = React.useState(formTypes.new), formTypeControl = _k[0], setFormTypeControl = _k[1];
    var _l = React.useState([]), menuListItems = _l[0], setMenuListItems = _l[1];
    var _m = React.useState(""), selectedListTitlle = _m[0], setSelectedListTitle = _m[1];
    var _o = React.useState(false), isLoading = _o[0], calendarIsLoading = _o[1];
    var _p = React.useState([]), eventSourcesArray = _p[0], setEvents = _p[1];
    var _q = React.useState({
        start: moment().startOf("month").toDate(),
        end: moment().endOf("month").toDate(),
    }), viewDateRange = _q[0], setViewDateRange = _q[1];
    var _dataService = new spservices(props.context);
    React.useEffect(function () {
        function fetchEvents() {
            var myEvents = [];
            var promises = props.selectedCalendarLists.map(function (calendarData) {
                return _dataService
                    .getEvents(escape(calendarData.SiteUrl), escape(calendarData.CalendarListTitle), viewDateRange.start, viewDateRange.end)
                    .then(function (eventsData) {
                    myEvents.push({
                        id: calendarData.CalendarTitle,
                        events: eventsData,
                        color: calendarData.Color
                    });
                })
                    .catch(function (error) {
                    console.error(error);
                });
            });
            return Promise.all(promises).then(function () { return myEvents; }); // Wait for all promises to resolve
        }
        fetchEvents()
            .then(function (myEvents) {
            console.log(myEvents);
            setEvents(myEvents);
            calendarIsLoading(false);
        })
            .catch(function (error) {
            console.error(error);
        });
    }, [viewDateRange, isLoading]);
    var handleDateSet = function (info) {
        var view = info.view, start = info.start, end = info.end;
        console.log("Visible range in ".concat(view.type, " view: ").concat(start, " - ").concat(end));
        setViewDateRange({ start: start, end: end });
        // calendarIsLoading(true);
    };
    var theme = getTheme();
    var calendarComponentRef = React.createRef();
    var navigateCalendar = function (date) {
        var calendarApi = calendarComponentRef.current.getApi();
        console.log(calendarApi.getEventSources());
        calendarApi.gotoDate(date);
    };
    return (React.createElement("div", { className: styles.spAggregatedCalendar },
        React.createElement(Stack, { tokens: containerStackTokens },
            React.createElement("hr", { color: theme.palette.themePrimary }),
            React.createElement(Stack, { tokens: verticalGapStackTokens },
                React.createElement(Stack, { horizontalAlign: "end" },
                    React.createElement(CommandBar, { items: [
                            {
                                key: "newItem",
                                text: "New Event",
                                cacheKey: "newItemCache",
                                iconProps: { iconName: "Add" },
                                subMenuProps: {
                                    items: menuItems,
                                },
                            },
                            {
                                key: "exportItems",
                                text: "Export Events",
                                cacheKey: "exportItemCache",
                                iconProps: { iconName: "ExcelDocument" },
                                subMenuProps: {
                                    items: menuItems,
                                },
                            },
                        ], 
                        // overflowButtonProps={overflowProps}
                        ariaLabel: "Calendar Commands", primaryGroupAriaLabel: "Email actions", farItemsGroupAriaLabel: "More actions" })),
                React.createElement(Stack, { horizontal: true, horizontalAlign: "space-between" },
                    React.createElement(Stack, null,
                        React.createElement(Calendar, { showMonthPickerAsOverlay: true, highlightSelectedMonth: true, showGoToToday: true, onSelectDate: function (date) { return navigateCalendar(date); }, 
                            // value={selectedDate}
                            // Calendar uses English strings by default. For localized apps, you must override this prop.
                            strings: defaultCalendarStrings }),
                        React.createElement(Stack, { horizontal: true, horizontalAlign: "start" },
                            " ",
                            props.showLegend && (React.createElement("div", null,
                                React.createElement("p", null, "Calendars in View:"),
                                React.createElement("div", { className: styles.legend },
                                    React.createElement(SpCalendarLegend, { selectedCalendarLists: props.selectedCalendarLists })))))),
                    React.createElement("div", null, isLoading ? React.createElement(Spinner, { size: SpinnerSize.large, label: strings.LoadingEventsLabel }) : React.createElement(FullCalendar, { plugins: [
                            timeGridPlugin,
                            dayGridPlugin,
                            rrulePlugin,
                            listPlugin,
                        ], initialView: props.defaultView, headerToolbar: {
                            center: "title",
                            start: "timeGridDay,timeGridWeek,dayGridMonth,listMonth",
                            end: "prev,next,today",
                        }, businessHours: {
                            // days of week. an array of zero-based day of week integers (0=Sunday)
                            daysOfWeek: [1, 2, 3, 4, 5],
                            startTime: "08:00",
                            endTime: "17:00", // an end time (6pm in this example)
                        }, initialDate: new Date(), navLinks: true, editable: true, aspectRatio: 2, ref: calendarComponentRef, datesSet: handleDateSet, 
                        // eventLimit = {3}
                        fixedWeekCount: false, 
                        // eventClick={this.eventClickHandler}
                        eventSources: eventSourcesArray })))))));
};
//# sourceMappingURL=SpAggregatedCalendar.js.map