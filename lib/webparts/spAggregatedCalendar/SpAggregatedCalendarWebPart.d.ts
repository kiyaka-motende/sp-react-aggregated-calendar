import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SelectedCalendar } from './models/ISelectedCalendar';
/**
 * Interface for the Aggregated Calendar Webpart Class Properties
 *
 * @export
 * @interface ISpAggregatedCalendarWebPartProps
 */
export interface ISpAggregatedCalendarWebPartProps {
    color: string;
    header: string;
    calendarList: SelectedCalendar[];
    dateFormat: string;
    showLegend: boolean;
    defaultView: string;
}
export default class SpAggregatedCalendarWebPart extends BaseClientSideWebPart<ISpAggregatedCalendarWebPartProps> {
    private _isDarkTheme;
    private _environmentMessage;
    private availableViews;
    private timeFormat;
    render(): void;
    protected onInit(): Promise<void>;
    private _getEnvironmentMessage;
    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=SpAggregatedCalendarWebPart.d.ts.map