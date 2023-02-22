var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
import * as React from "react";
import * as ReactDom from "react-dom";
import { Log, Version } from "@microsoft/sp-core-library";
import { PropertyPaneDropdown, PropertyPaneToggle, } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as strings from "SpAggregatedCalendarWebPartStrings";
import { SpAggregatedCalendar } from "./components/calendar_components/mainCalendar/SpAggregatedCalendar";
import { PropertyFieldCollectionData, CustomCollectionFieldType, } from "@pnp/spfx-property-controls/lib/PropertyFieldCollectionData";
import { MessageBarType } from "@fluentui/react";
import { MessageComponent, } from "../shared/components/MessageComponent";
var SpAggregatedCalendarWebPart = /** @class */ (function (_super) {
    __extends(SpAggregatedCalendarWebPart, _super);
    function SpAggregatedCalendarWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._environmentMessage = "";
        _this.availableViews = require("../shared/availableViews.json");
        _this.timeFormat = require("../shared/timeFormat.json");
        return _this;
    }
    SpAggregatedCalendarWebPart.prototype.render = function () {
        Log.verbose("render()", "Inside Render", this.context.serviceScope);
        if (this.needsConfiguration()) {
            Log.warn("render()", "Webpart not configured", this.context.serviceScope);
            this.renderMessage(strings.WebPartNotConfigured, MessageBarType.error, true);
        }
        else {
            Log.info("render()", "Webpart configuration not needed", this.context.serviceScope);
            var element = React.createElement(SpAggregatedCalendar, {
                header: this.properties.header,
                selectedCalendarLists: this.properties.calendarList,
                context: this.context,
                domElement: this.domElement,
                dateFormat: this.properties.dateFormat,
                showLegend: this.properties.showLegend,
                defaultView: this.properties.defaultView,
            });
            ReactDom.render(element, this.domElement);
        }
    };
    SpAggregatedCalendarWebPart.prototype.onInit = function () {
        var _this = this;
        return this._getEnvironmentMessage().then(function (message) {
            _this._environmentMessage = message;
        });
    };
    SpAggregatedCalendarWebPart.prototype._getEnvironmentMessage = function () {
        var _this = this;
        if (!!this.context.sdks.microsoftTeams) {
            // running in Teams, office.com or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app
                .getContext()
                .then(function (context) {
                var environmentMessage = "";
                switch (context.app.host.name) {
                    case "Office": // running in Office
                        environmentMessage = _this.context.isServedFromLocalhost
                            ? strings.AppLocalEnvironmentOffice
                            : strings.AppOfficeEnvironment;
                        break;
                    case "Outlook": // running in Outlook
                        environmentMessage = _this.context.isServedFromLocalhost
                            ? strings.AppLocalEnvironmentOutlook
                            : strings.AppOutlookEnvironment;
                        break;
                    case "Teams": // running in Teams
                        environmentMessage = _this.context.isServedFromLocalhost
                            ? strings.AppLocalEnvironmentTeams
                            : strings.AppTeamsTabEnvironment;
                        break;
                    default:
                        throw new Error("Unknown host");
                }
                return environmentMessage;
            });
        }
        return Promise.resolve(this.context.isServedFromLocalhost
            ? strings.AppLocalEnvironmentSharePoint
            : strings.AppSharePointEnvironment);
    };
    SpAggregatedCalendarWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors;
        if (semanticColors) {
            this.domElement.style.setProperty("--bodyText", semanticColors.bodyText || null);
            this.domElement.style.setProperty("--link", semanticColors.link || null);
            this.domElement.style.setProperty("--linkHovered", semanticColors.linkHovered || null);
        }
    };
    SpAggregatedCalendarWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(SpAggregatedCalendarWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse("1.0");
        },
        enumerable: false,
        configurable: true
    });
    SpAggregatedCalendarWebPart.prototype.getPropertyPaneConfiguration = function () {
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
                                PropertyFieldCollectionData("calendarList", {
                                    key: "calendarList",
                                    label: strings.SelectCalendarLabel,
                                    value: this.properties.calendarList,
                                    manageBtnLabel: "Manage Calendar",
                                    panelHeader: "Add Calendars",
                                    fields: [
                                        {
                                            id: "CalendarTitle",
                                            title: "Calendar Title",
                                            required: true,
                                            type: CustomCollectionFieldType.string,
                                        },
                                        {
                                            id: "SiteUrl",
                                            title: "Site Url",
                                            required: true,
                                            type: CustomCollectionFieldType.string,
                                        },
                                        {
                                            id: "CalendarListTitle",
                                            title: "Calendar List Title",
                                            required: true,
                                            type: CustomCollectionFieldType.string,
                                        },
                                        {
                                            id: "Color",
                                            title: "Color",
                                            required: false,
                                            type: CustomCollectionFieldType.color,
                                        },
                                    ],
                                    disabled: false,
                                }),
                                PropertyPaneDropdown("dateFormat", {
                                    label: strings.SelectDateFormatFieldLabel,
                                    selectedKey: "MMMM Do YYYY, h: mm a",
                                    options: this.timeFormat,
                                }),
                                PropertyPaneToggle("showLegend", {
                                    label: strings.ShowLegendFieldLabel,
                                    onText: strings.OnTextFieldLabel,
                                    offText: strings.OffTextFieldLabel,
                                    checked: false,
                                }),
                                PropertyPaneDropdown("defaultView", {
                                    label: strings.DefaultView,
                                    selectedKey: "dayGridMonth",
                                    options: this.availableViews,
                                }),
                            ],
                        },
                    ],
                },
            ],
        };
    };
    /**
     * Check whether Aggregated Calendar needs configuration
     * or not
     * @private
     * @returns {boolean}
     * @memberof ReactAggregatedCalendarWebPart
     */
    SpAggregatedCalendarWebPart.prototype.needsConfiguration = function () {
        Log.verbose("needsConfiguration()", "calendarList : " + this.properties.calendarList, this.context.serviceScope);
        return (this.properties.calendarList === null ||
            this.properties.calendarList === undefined ||
            this.properties.calendarList.length === 0);
    };
    /**
     * Render Message method to render the message component
     *
     * @private
     * @param {string} statusMessage
     * @param {MessageBarType} statusMessageType
     * @param {boolean} display
     * @memberof ReactAggregatedCalendarWebPart
     */
    SpAggregatedCalendarWebPart.prototype.renderMessage = function (statusMessage, statusMessageType, display) {
        Log.verbose("renderMessage()", "Rendering Message " + statusMessage + " of type " + statusMessageType, this.context.serviceScope);
        var messageElement = React.createElement(MessageComponent, {
            Message: statusMessage,
            Type: statusMessageType,
            Display: display,
        });
        ReactDom.render(messageElement, this.domElement);
    };
    return SpAggregatedCalendarWebPart;
}(BaseClientSideWebPart));
export default SpAggregatedCalendarWebPart;
//# sourceMappingURL=SpAggregatedCalendarWebPart.js.map