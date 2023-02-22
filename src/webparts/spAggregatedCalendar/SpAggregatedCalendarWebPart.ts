import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpAggregatedCalendarWebPartStrings';
import {SpAggregatedCalendar} from './components/calendar_components/mainCalendar/SpAggregatedCalendar';
import { ISpAggregatedCalendarProps } from './components/calendar_components/mainCalendar/ISpAggregatedCalendarProps';
import { SelectedCalendar } from './models/ISelectedCalendar';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { IDropdownOption } from '@fluentui/react';


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

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private availableViews: IDropdownOption[] = require("../shared/availableViews.json");
  private timeFormat: IDropdownOption[] = require("../shared/timeFormat.json");
  
  public render(): void {
    const element: React.ReactElement<ISpAggregatedCalendarProps> = React.createElement(
      SpAggregatedCalendar,
      {
        header: this.properties.header,
        selectedCalendarLists: this.properties.calendarList,
        context: this.context,
        domElement: this.domElement,
        dateFormat: this.properties.dateFormat,
        showLegend: this.properties.showLegend,
        defaultView: this.properties.defaultView
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          // header: {
          //   description: strings.PropertyPaneDescription
          // },
          groups: [
            {
              // groupName: strings.BasicGroupName,
              groupFields: [
                // PropertyPaneTextField('header', {
                //   label: strings.HeaderFieldLabel
                // }),
                
                PropertyFieldCollectionData('calendarList', {
                  key: 'calendarList',
                  label: strings.SelectCalendarLabel,
                  value: this.properties.calendarList,
                  manageBtnLabel: 'Manage Calendar',
                  panelHeader:'Add Calendars',
                  fields: [
                    { id: 'CalendarTitle', title: 'Calendar Title', required: true, type: CustomCollectionFieldType.string },
                    { id: 'SiteUrl', title: 'Site Url', required: true, type: CustomCollectionFieldType.string },
                    {
                      id: 'CalendarListTitle', title: 'Calendar List Title', required: true,
                      type: CustomCollectionFieldType.string
                    },
                    { id: 'Color', title: 'Color', required: false, type: CustomCollectionFieldType.color }
                  ],
                  disabled: false
                }),
                
                PropertyPaneDropdown('dateFormat', {
                  label: strings.SelectDateFormatFieldLabel,
                  selectedKey: "MMMM Do YYYY, h: mm a",
                  options: this.timeFormat
                }),
                PropertyPaneToggle('showLegend', {
                  label: strings.ShowLegendFieldLabel,
                  onText: strings.OnTextFieldLabel,
                  offText: strings.OffTextFieldLabel,
                  checked: false
                }),
                PropertyPaneDropdown('defaultView',{
                  label: strings.DefaultView,
                  selectedKey: "dayGridMonth",
                  options: this.availableViews
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
