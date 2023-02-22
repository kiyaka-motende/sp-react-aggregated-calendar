import * as React from "react";
import styles from "./SpAggregatedCalendar.module.scss";
import * as moment from "moment";
export var formTypes;
(function (formTypes) {
    formTypes[formTypes["new"] = 1] = "new";
    formTypes[formTypes["edit"] = 2] = "edit";
    formTypes[formTypes["delete"] = 3] = "delete";
})(formTypes || (formTypes = {}));
export var SpAggregatedCalendar = function (props) {
    var _a = React.useState(false), isCalendarFiltered = _a[0], setisClaendarFiltered = _a[1];
    var _b = React.useState(""), selectedCalendar = _b[0], setSelectedCalendar = _b[1];
    var _c = React.useState(false), isCalloutVisible = _c[0], setCalloutVisible = _c[1];
    var _d = React.useState([]), selectedCalendarList = _d[0], setSelectedCalendarList = _d[1];
    var _e = React.useState([]), calendarEvents = _e[0], setCalendarEvents = _e[1];
    var _f = React.useState({
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
    }), selectedEvent = _f[0], setSelectedEvent = _f[1];
    var _g = React.useState([]), menuItems = _g[0], setMenuItems = _g[1];
    var _h = React.useState([]), filteredLists = _h[0], setFilteredList = _h[1];
    var _j = React.useState(false), isModalOpen = _j[0], setIsModalOpen = _j[1];
    var _k = React.useState(true), isDialogHidden = _k[0], setIsDialogHidden = _k[1];
    var _l = React.useState(''), selectedListId = _l[0], setSelectedListId = _l[1];
    var _m = React.useState(formTypes.new), formTypeControl = _m[0], setFormTypeControl = _m[1];
    var _o = React.useState([]), menuListItems = _o[0], setMenuListItems = _o[1];
    var _p = React.useState(''), selectedListTitlle = _p[0], setSelectedListTitle = _p[1];
    return (React.createElement("div", { className: styles.spAggregatedCalendar }));
};
//# sourceMappingURL=SpAggregatedCalendar.js.map