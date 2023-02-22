/* eslint-disable react/self-closing-comp */
import * as React from "react";
import { ISpCalendarLegendProps } from "./ISpCalendarLegendProps";
import styles from "./SpCalendarLegend.module.scss";

export const SpCalendarLegend: React.FunctionComponent<
  ISpCalendarLegendProps
> = (props: ISpCalendarLegendProps) => {
  const [isCalendarFiltered, setisCalendarFiltered] = React.useState(false);
  const [selectedCalendar, setSelectedCalendar] = React.useState("");
  let calendarLegend: JSX.Element[] = [];

  // Render the Legend for the Calendar Events
  calendarLegend = props.selectedCalendarLists.map((calendar) => {
    const calendarLegendColor = {
      backgroundColor: `${calendar.Color}`,
    };
    console.log(calendarLegendColor);
    let outerClass: string;
    if (!isCalendarFiltered) {
      outerClass = styles.selected;
    } else if (
      isCalendarFiltered &&
      selectedCalendar === calendar.CalendarListTitle
    ) {
      outerClass = styles.selected;
    } else outerClass = styles.washout;

    return (
      // eslint-disable-next-line react/jsx-key
      <div
        className={`${styles.outerLegendDiv} ${outerClass}`}
        title={calendar.CalendarTitle}
      >
        <div
          className={styles.innerLegendDiv}
          style={calendarLegendColor}
          onClick={(e) => {
            setisCalendarFiltered(!isCalendarFiltered);
            setSelectedCalendar(calendar.CalendarListTitle);
          }}
        ></div>
        {calendar.CalendarTitle}
      </div>
    );
  });
  return <span>{calendarLegend}</span>;
};
