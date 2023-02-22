import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SelectedCalendar } from "../../models/ISelectedCalendar";
export interface ISpAggregatedCalendarProps {
    header: string;
    selectedCalendarLists: SelectedCalendar[];
    context: WebPartContext;
    domElement: HTMLElement;
    dateFormat: string;
    showLegend: boolean;
    defaultView: string;
}
//# sourceMappingURL=ISpAggregatedCalendarProps.d.ts.map