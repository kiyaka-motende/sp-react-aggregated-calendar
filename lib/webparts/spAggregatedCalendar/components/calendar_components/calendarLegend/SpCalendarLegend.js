/* eslint-disable react/self-closing-comp */
import * as React from "react";
import styles from "./SpCalendarLegend.module.scss";
export var SpCalendarLegend = function (props) {
    var _a = React.useState(false), isCalendarFiltered = _a[0], setisCalendarFiltered = _a[1];
    var _b = React.useState(""), selectedCalendar = _b[0], setSelectedCalendar = _b[1];
    var calendarLegend = [];
    // Render the Legend for the Calendar Events
    calendarLegend = props.selectedCalendarLists.map(function (calendar) {
        var calendarLegendColor = {
            backgroundColor: "".concat(calendar.Color),
        };
        var outerClass;
        if (!isCalendarFiltered) {
            outerClass = styles.selected;
        }
        else if (isCalendarFiltered &&
            selectedCalendar === calendar.CalendarListTitle) {
            outerClass = styles.selected;
        }
        else
            outerClass = styles.washout;
        return (
        // eslint-disable-next-line react/jsx-key
        React.createElement("div", { className: "".concat(styles.outerLegendDiv, " ").concat(outerClass), title: calendar.CalendarTitle },
            React.createElement("div", { className: "".concat(styles.innerLegendDiv), style: calendarLegendColor, onClick: function (e) {
                    setisCalendarFiltered(!isCalendarFiltered);
                    setSelectedCalendar(calendar.CalendarListTitle);
                } }),
            calendar.CalendarTitle));
    });
    return (React.createElement("div", { className: styles.calendarLegend }, calendarLegend));
};
//# sourceMappingURL=SpCalendarLegend.js.map