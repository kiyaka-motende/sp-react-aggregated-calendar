var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { sp, Web } from "@pnp/sp";
import * as moment from "moment";
import { RRule } from "rrule";
import parseRecurrentEvent from "./parseRecurrentEvent";
// import { graph} from "@pnp/graph";
// Class Services
var spservices = /** @class */ (function () {
    function spservices(context) {
        this.context = context;
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
    spservices.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () { return __generator(this, function (_a) {
            return [2 /*return*/];
        }); });
    };
    /**
     *
     * @private
     * @returns {Promise<string>}
     * @memberof spservices
     */
    spservices.prototype.getLocalTime = function (date) {
        return __awaiter(this, void 0, void 0, function () {
            var localTime, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, sp.web.regionalSettings.timeZone.utcToLocalTime(date)];
                    case 1:
                        localTime = _a.sent();
                        return [2 /*return*/, localTime];
                    case 2:
                        error_1 = _a.sent();
                        return [2 /*return*/, Promise.reject(error_1)];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     *
     * @private
     * @returns {Promise<string>}
     * @memberof spservices
     */
    spservices.prototype.getUtcTime = function (date) {
        return __awaiter(this, void 0, void 0, function () {
            var utcTime, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, sp.web.regionalSettings.timeZone.localTimeToUTC(date)];
                    case 1:
                        utcTime = _a.sent();
                        return [2 /*return*/, utcTime];
                    case 2:
                        error_2 = _a.sent();
                        return [2 /*return*/, Promise.reject(error_2)];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     *
     * @param {string} siteUrl
     * @param {string} listId
     * @param {string} fieldInternalName
     * @returns {Promise<{ key: string, text: string }[]>}
     * @memberof spservices
     */
    spservices.prototype.getChoiceFieldOptions = function (siteUrl, listId, fieldInternalName) {
        return __awaiter(this, void 0, void 0, function () {
            var fieldOptions, web, results, _i, _a, option, error_3;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        fieldOptions = [];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        web = new Web(siteUrl);
                        return [4 /*yield*/, web.lists
                                .getById(listId)
                                .fields.getByInternalNameOrTitle(fieldInternalName)
                                .select("Title", "InternalName", "Choices")
                                .get()];
                    case 2:
                        results = _b.sent();
                        if (results && results.Choices.length > 0) {
                            for (_i = 0, _a = results.Choices; _i < _a.length; _i++) {
                                option = _a[_i];
                                fieldOptions.push({
                                    key: option,
                                    text: option,
                                });
                            }
                        }
                        return [3 /*break*/, 4];
                    case 3:
                        error_3 = _b.sent();
                        return [2 /*return*/, Promise.reject(error_3)];
                    case 4: return [2 /*return*/, fieldOptions];
                }
            });
        });
    };
    /**
     *
     * @param {string} siteUrl
     * @param {string} listId
     * @param {Date} eventStartDate
     * @param {Date} eventEndDate
     * @returns {Promise< IEventData[]>}
     * @memberof spservices
     */
    spservices.prototype.getEvents = function (siteUrl, listTitle, eventStartDate, eventEndDate) {
        return __awaiter(this, void 0, void 0, function () {
            var events, parseEvt, web, web_1, results_1, event_1, mapEvents, error_4;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        events = [];
                        if (!siteUrl) {
                            return [2 /*return*/, []];
                        }
                        parseEvt = new parseRecurrentEvent();
                        web = new Web(siteUrl);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 9, , 10]);
                        web_1 = new Web(siteUrl);
                        return [4 /*yield*/, web_1.lists
                                .getByTitle(listTitle)
                                .usingCaching()
                                .renderListDataAsStream({
                                DatesInUtc: true,
                                ViewXml: "<View><ViewFields><FieldRef Name='RecurrenceData'/><FieldRef Name='Duration'/><FieldRef Name='Author'/><FieldRef Name='Category'/><FieldRef Name='Description'/><FieldRef Name='ParticipantsPicker'/><FieldRef Name='Geolocation'/><FieldRef Name='ID'/><FieldRef Name='EndDate'/><FieldRef Name='EventDate'/><FieldRef Name='ID'/><FieldRef Name='Location'/><FieldRef Name='Title'/><FieldRef Name='fAllDayEvent'/><FieldRef Name='EventType'/><FieldRef Name='UID' /><FieldRef Name='fRecurrence' /></ViewFields>\n          <Query>\n          <Where>\n            <And>\n              <Geq>\n                <FieldRef Name='EventDate' />\n                <Value IncludeTimeValue='false' Type='DateTime'>".concat(moment(eventStartDate).format('YYYY-MM-DD'), "</Value>\n              </Geq>\n              <Leq>\n                <FieldRef Name='EventDate' />\n                <Value IncludeTimeValue='false' Type='DateTime'>").concat(moment(eventEndDate).format('YYYY-MM-DD'), "</Value>\n              </Leq>\n              </And>\n          </Where>\n          </Query>      \n          <RowLimit Paged=\"FALSE\">2000</RowLimit>\n          </View>"),
                            })];
                    case 2:
                        results_1 = _a.sent();
                        if (!(results_1 && results_1.Row.length > 0)) return [3 /*break*/, 8];
                        event_1 = "";
                        mapEvents = function () { return __awaiter(_this, void 0, void 0, function () {
                            var _i, _a, eventDate, endDate, initialsArray, initials, attendees, first, last, geo, geolocation, isAllDayEvent, _b, _c, attendee, _d, _e, _f, _g, _h, _j, _k, _l;
                            var _m;
                            return __generator(this, function (_o) {
                                switch (_o.label) {
                                    case 0:
                                        _i = 0, _a = results_1.Row;
                                        _o.label = 1;
                                    case 1:
                                        if (!(_i < _a.length)) return [3 /*break*/, 13];
                                        event_1 = _a[_i];
                                        return [4 /*yield*/, this.getLocalTime(event_1.EventDate)];
                                    case 2:
                                        eventDate = _o.sent();
                                        return [4 /*yield*/, this.getLocalTime(event_1.EndDate)];
                                    case 3:
                                        endDate = _o.sent();
                                        initialsArray = event_1.Author[0].title.split(" ");
                                        initials = initialsArray[0].charAt(0) +
                                            initialsArray[initialsArray.length - 1].charAt(0);
                                        attendees = [];
                                        first = event_1.Geolocation.indexOf("(") + 1;
                                        last = event_1.Geolocation.indexOf(")");
                                        geo = event_1.Geolocation.substring(first, last);
                                        geolocation = geo.split(" ");
                                        isAllDayEvent = event_1["fAllDayEvent.value"] === "1";
                                        for (_b = 0, _c = event_1.ParticipantsPicker; _b < _c.length; _b++) {
                                            attendee = _c[_b];
                                            attendees.push(parseInt(attendee.id));
                                        }
                                        _e = (_d = events).push;
                                        _m = {
                                            Id: event_1.ID,
                                            ID: event_1.ID,
                                            EventType: event_1.EventType
                                        };
                                        return [4 /*yield*/, parseEvt.deCodeHtmlEntities(event_1.Title)];
                                    case 4:
                                        _m.title = _o.sent(),
                                            _m.Description = event_1.Description,
                                            _m.start = isAllDayEvent
                                                ? new Date(event_1.EventDate.slice(0, -1))
                                                : new Date(eventDate),
                                            _m.end = isAllDayEvent
                                                ? new Date(event_1.EndDate.slice(0, -1))
                                                : new Date(endDate),
                                            _m.location = event_1.Location,
                                            _m.ownerEmail = event_1.Author[0].email,
                                            // ownerPhoto: userPictureUrl ?
                                            //   `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${event.Author[0].email}&UA=0&size=HR96x96` : '',
                                            _m.ownerInitial = initials,
                                            // color: CategoryColorValue.length > 0 ? CategoryColorValue[0].color : '#1a75ff', // blue default
                                            _m.ownerName = event_1.Author[0].title,
                                            _m.attendes = attendees,
                                            _m.fAllDayEvent = isAllDayEvent,
                                            _m.geolocation = {
                                                Longitude: parseFloat(geolocation[0]),
                                                Latitude: parseFloat(geolocation[1]),
                                            },
                                            _m.Category = event_1.Category,
                                            _m.Duration = event_1.Duration;
                                        if (!event_1.fRecurrence) return [3 /*break*/, 10];
                                        _h = (_g = RRule).fromText;
                                        if (!(event_1.EventType === "4" && event_1.MasterSeriesItemID !== "")) return [3 /*break*/, 6];
                                        return [4 /*yield*/, parseEvt.deCodeHtmlEntities(event_1.RecurrenceData)];
                                    case 5:
                                        _j = _o.sent();
                                        return [3 /*break*/, 9];
                                    case 6:
                                        _l = (_k = parseEvt).returnExceptionRecurrenceInfo;
                                        return [4 /*yield*/, parseEvt.deCodeHtmlEntities(event_1.RecurrenceData)];
                                    case 7: return [4 /*yield*/, _l.apply(_k, [_o.sent()])];
                                    case 8:
                                        _j = _o.sent();
                                        _o.label = 9;
                                    case 9:
                                        _f = _h.apply(_g, [_j]).options;
                                        return [3 /*break*/, 11];
                                    case 10:
                                        _f = null;
                                        _o.label = 11;
                                    case 11:
                                        _e.apply(_d, [(_m.rrule = _f,
                                                _m.fRecurrence = event_1.fRecurrence,
                                                _m.RecurrenceID = event_1.RecurrenceID ? event_1.RecurrenceID : undefined,
                                                _m.MasterSeriesItemID = event_1.MasterSeriesItemID,
                                                _m.UID = event_1.UID.replace("{", "").replace("}", ""),
                                                _m)]);
                                        _o.label = 12;
                                    case 12:
                                        _i++;
                                        return [3 /*break*/, 1];
                                    case 13: return [2 /*return*/, true];
                                }
                            });
                        }); };
                        if (!window.localStorage.getItem("eventResult")) return [3 /*break*/, 6];
                        if (!(window.localStorage.getItem("eventResult") ===
                            JSON.stringify(results_1))) return [3 /*break*/, 3];
                        //No update needed use current savedEvents
                        events = JSON.parse(window.localStorage.getItem("calendarEventsWithLocalTime"));
                        return [3 /*break*/, 5];
                    case 3:
                        //update local storage
                        window.localStorage.setItem("eventResult", JSON.stringify(results_1));
                        return [4 /*yield*/, mapEvents()];
                    case 4:
                        //when they are not equal then we loop through the results and maps them to IEventData
                        /* tslint:disable:no-unused-expression */
                        (_a.sent())
                            ? window.localStorage.setItem("calendarEventsWithLocalTime", JSON.stringify(events))
                            : null;
                        _a.label = 5;
                    case 5: return [3 /*break*/, 8];
                    case 6:
                        //if there is no local storage of the events we create them
                        window.localStorage.setItem("eventResult", JSON.stringify(results_1));
                        return [4 /*yield*/, mapEvents()];
                    case 7:
                        //we also needs to map through the events the first time and save the mapped version to local storage
                        (_a.sent())
                            ? window.localStorage.setItem("calendarEventsWithLocalTime", JSON.stringify(events))
                            : null;
                        _a.label = 8;
                    case 8: 
                    //   events = parseEvt.parseEvents(events, null, null);
                    // Return Data
                    return [2 /*return*/, events];
                    case 9:
                        error_4 = _a.sent();
                        console.dir(error_4);
                        return [2 /*return*/, Promise.reject(error_4)];
                    case 10: return [2 /*return*/];
                }
            });
        });
    };
    return spservices;
}());
export default spservices;
//# sourceMappingURL=spservices.js.map