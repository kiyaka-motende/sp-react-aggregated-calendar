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
/* eslint-disable dot-notation */
/* eslint-disable @typescript-eslint/no-var-requires */
// import { RRule } from "rrule";
import { parseString } from "xml2js";
import * as strings from "SpAggregatedCalendarWebPartStrings";
var format = require("string-format");
var DayPickerStrings = {
    months: [
        strings.January,
        strings.February,
        strings.March,
        strings.April,
        strings.May,
        strings.June,
        strings.July,
        strings.August,
        strings.September,
        strings.October,
        strings.November,
        strings.December,
    ],
    shortMonths: [
        strings.Jan,
        strings.Feb,
        strings.Mar,
        strings.Apr,
        strings.May,
        strings.Jun,
        strings.Jul,
        strings.Aug,
        strings.Sep,
        strings.Oct,
        strings.Nov,
        strings.Dez,
    ],
    days: [
        strings.Sunday,
        strings.Monday,
        strings.Tuesday,
        strings.Wednesday,
        strings.Thursday,
        strings.Friday,
        strings.Saturday,
    ],
    shortDays: [
        strings.ShortDay_S,
        strings.ShortDay_M,
        strings.ShortDay_T,
        strings.ShortDay_W,
        strings.ShortDay_Thursday,
        strings.ShortDay_Friday,
        strings.ShortDay_Sunday,
    ],
    goToToday: strings.GoToDay,
    prevMonthAriaLabel: strings.PrevMonth,
    nextMonthAriaLabel: strings.NextMonth,
    prevYearAriaLabel: strings.PrevYear,
    nextYearAriaLabel: strings.NextYear,
    closeButtonAriaLabel: strings.CloseDate,
    isRequiredErrorMessage: strings.IsRequired,
    invalidInputErrorMessage: strings.InvalidDateFormat,
};
var parseRecurrentEvent = /** @class */ (function () {
    function parseRecurrentEvent() {
    }
    parseRecurrentEvent.prototype.deCodeHtmlEntities = function (string) {
        return __awaiter(this, void 0, void 0, function () {
            var HtmlEntitiesMap, entityMap, key, entity, regex;
            return __generator(this, function (_a) {
                HtmlEntitiesMap = {
                    "'": "&#39;",
                    "<": "&lt;",
                    ">": "&gt;",
                    " ": "&nbsp;",
                    "¡": "&iexcl;",
                    "¢": "&cent;",
                    "£": "&pound;",
                    "¤": "&curren;",
                    "¥": "&yen;",
                    "¦": "&brvbar;",
                    "§": "&sect;",
                    "¨": "&uml;",
                    "©": "&copy;",
                    "ª": "&ordf;",
                    "«": "&laquo;",
                    "¬": "&not;",
                    "®": "&reg;",
                    "¯": "&macr;",
                    "°": "&deg;",
                    "±": "&plusmn;",
                    "²": "&sup2;",
                    "³": "&sup3;",
                    "´": "&acute;",
                    "µ": "&micro;",
                    "¶": "&para;",
                    "·": "&middot;",
                    "¸": "&cedil;",
                    "¹": "&sup1;",
                    "º": "&ordm;",
                    "»": "&raquo;",
                    "¼": "&frac14;",
                    "½": "&frac12;",
                    "¾": "&frac34;",
                    "¿": "&iquest;",
                    "À": "&Agrave;",
                    "Á": "&Aacute;",
                    "Â": "&Acirc;",
                    "Ã": "&Atilde;",
                    "Ä": "&Auml;",
                    "Å": "&Aring;",
                    "Æ": "&AElig;",
                    "Ç": "&Ccedil;",
                    "È": "&Egrave;",
                    "É": "&Eacute;",
                    "Ê": "&Ecirc;",
                    "Ë": "&Euml;",
                    "Ì": "&Igrave;",
                    "Í": "&Iacute;",
                    "Î": "&Icirc;",
                    "Ï": "&Iuml;",
                    "Ð": "&ETH;",
                    "Ñ": "&Ntilde;",
                    "Ò": "&Ograve;",
                    "Ó": "&Oacute;",
                    "Ô": "&Ocirc;",
                    "Õ": "&Otilde;",
                    "Ö": "&Ouml;",
                    "×": "&times;",
                    "Ø": "&Oslash;",
                    "Ù": "&Ugrave;",
                    "Ú": "&Uacute;",
                    "Û": "&Ucirc;",
                    "Ü": "&Uuml;",
                    "Ý": "&Yacute;",
                    "Þ": "&THORN;",
                    "ß": "&szlig;",
                    "à": "&agrave;",
                    "á": "&aacute;",
                    "â": "&acirc;",
                    "ã": "&atilde;",
                    "ä": "&auml;",
                    "å": "&aring;",
                    "æ": "&aelig;",
                    "ç": "&ccedil;",
                    "è": "&egrave;",
                    "é": "&eacute;",
                    "ê": "&ecirc;",
                    "ë": "&euml;",
                    "ì": "&igrave;",
                    "í": "&iacute;",
                    "î": "&icirc;",
                    "ï": "&iuml;",
                    "ð": "&eth;",
                    "ñ": "&ntilde;",
                    "ò": "&ograve;",
                    "ó": "&oacute;",
                    "ô": "&ocirc;",
                    "õ": "&otilde;",
                    "ö": "&ouml;",
                    "÷": "&divide;",
                    "ø": "&oslash;",
                    "ù": "&ugrave;",
                    "ú": "&uacute;",
                    "û": "&ucirc;",
                    "ü": "&uuml;",
                    "ý": "&yacute;",
                    "þ": "&thorn;",
                    "ÿ": "&yuml;",
                    "Œ": "&OElig;",
                    "œ": "&oelig;",
                    "Š": "&Scaron;",
                    "š": "&scaron;",
                    "Ÿ": "&Yuml;",
                    "ƒ": "&fnof;",
                    "ˆ": "&circ;",
                    "˜": "&tilde;",
                    "Α": "&Alpha;",
                    "Β": "&Beta;",
                    "Γ": "&Gamma;",
                    "Δ": "&Delta;",
                    "Ε": "&Epsilon;",
                    "Ζ": "&Zeta;",
                    "Η": "&Eta;",
                    "Θ": "&Theta;",
                    "Ι": "&Iota;",
                    "Κ": "&Kappa;",
                    "Λ": "&Lambda;",
                    "Μ": "&Mu;",
                    "Ν": "&Nu;",
                    "Ξ": "&Xi;",
                    "Ο": "&Omicron;",
                    "Π": "&Pi;",
                    "Ρ": "&Rho;",
                    "Σ": "&Sigma;",
                    "Τ": "&Tau;",
                    "Υ": "&Upsilon;",
                    "Φ": "&Phi;",
                    "Χ": "&Chi;",
                    "Ψ": "&Psi;",
                    "Ω": "&Omega;",
                    "α": "&alpha;",
                    "β": "&beta;",
                    "γ": "&gamma;",
                    "δ": "&delta;",
                    "ε": "&epsilon;",
                    "ζ": "&zeta;",
                    "η": "&eta;",
                    "θ": "&theta;",
                    "ι": "&iota;",
                    "κ": "&kappa;",
                    "λ": "&lambda;",
                    "μ": "&mu;",
                    "ν": "&nu;",
                    "ξ": "&xi;",
                    "ο": "&omicron;",
                    "π": "&pi;",
                    "ρ": "&rho;",
                    "ς": "&sigmaf;",
                    "σ": "&sigma;",
                    "τ": "&tau;",
                    "υ": "&upsilon;",
                    "φ": "&phi;",
                    "χ": "&chi;",
                    "ψ": "&psi;",
                    "ω": "&omega;",
                    "ϑ": "&thetasym;",
                    "ϒ": "&Upsih;",
                    "ϖ": "&piv;",
                    "–": "&ndash;",
                    "—": "&mdash;",
                    "‘": "&lsquo;",
                    "’": "&rsquo;",
                    "‚": "&sbquo;",
                    "“": "&ldquo;",
                    "”": "&rdquo;",
                    "„": "&bdquo;",
                    "†": "&dagger;",
                    "‡": "&Dagger;",
                    "•": "&bull;",
                    "…": "&hellip;",
                    "‰": "&permil;",
                    "′": "&prime;",
                    "″": "&Prime;",
                    "‹": "&lsaquo;",
                    "›": "&rsaquo;",
                    "‾": "&oline;",
                    "⁄": "&frasl;",
                    "€": "&euro;",
                    "ℑ": "&image;",
                    "℘": "&weierp;",
                    "ℜ": "&real;",
                    "™": "&trade;",
                    "ℵ": "&alefsym;",
                    "←": "&larr;",
                    "↑": "&uarr;",
                    "→": "&rarr;",
                    "↓": "&darr;",
                    "↔": "&harr;",
                    "↵": "&crarr;",
                    "⇐": "&lArr;",
                    "⇑": "&UArr;",
                    "⇒": "&rArr;",
                    "⇓": "&dArr;",
                    "⇔": "&hArr;",
                    "∀": "&forall;",
                    "∂": "&part;",
                    "∃": "&exist;",
                    "∅": "&empty;",
                    "∇": "&nabla;",
                    "∈": "&isin;",
                    "∉": "&notin;",
                    "∋": "&ni;",
                    "∏": "&prod;",
                    "∑": "&sum;",
                    "−": "&minus;",
                    "∗": "&lowast;",
                    "√": "&radic;",
                    "∝": "&prop;",
                    "∞": "&infin;",
                    "∠": "&ang;",
                    "∧": "&and;",
                    "∨": "&or;",
                    "∩": "&cap;",
                    "∪": "&cup;",
                    "∫": "&int;",
                    "∴": "&there4;",
                    "∼": "&sim;",
                    "≅": "&cong;",
                    "≈": "&asymp;",
                    "≠": "&ne;",
                    "≡": "&equiv;",
                    "≤": "&le;",
                    "≥": "&ge;",
                    "⊂": "&sub;",
                    "⊃": "&sup;",
                    "⊄": "&nsub;",
                    "⊆": "&sube;",
                    "⊇": "&supe;",
                    "⊕": "&oplus;",
                    "⊗": "&otimes;",
                    "⊥": "&perp;",
                    "⋅": "&sdot;",
                    "⌈": "&lceil;",
                    "⌉": "&rceil;",
                    "⌊": "&lfloor;",
                    "⌋": "&rfloor;",
                    "⟨": "&lang;",
                    "⟩": "&rang;",
                    "◊": "&loz;",
                    "♠": "&spades;",
                    "♣": "&clubs;",
                    "♥": "&hearts;",
                    "♦": "&diams;"
                };
                entityMap = HtmlEntitiesMap;
                for (key in entityMap) {
                    if (Object.prototype.hasOwnProperty.call(entityMap, key)) {
                        entity = entityMap[key];
                        regex = new RegExp(entity, 'g');
                        string = string.replace(regex, key);
                    }
                }
                string = string.replace(/&quot;/g, '"');
                string = string.replace(/&amp;/g, '&');
                return [2 /*return*/, string];
            });
        });
    };
    /**
     *
     *
     * @private
     * @param {string} recurrenceData
     * @memberof Event
     */
    parseRecurrentEvent.prototype.returnExceptionRecurrenceInfo = function (recurrenceData) {
        return __awaiter(this, void 0, void 0, function () {
            var promise, recurrenceInfo, keys, recurrenceTypes, freq, _i, keys_1, key, rule;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        promise = new Promise(function (resolve, reject) {
                            parseString(recurrenceData, function (err, result) {
                                if (err) {
                                    reject(err);
                                }
                                resolve(result);
                            });
                        });
                        return [4 /*yield*/, promise];
                    case 1:
                        recurrenceInfo = _a.sent();
                        if (recurrenceInfo !== null) {
                            keys = Object.keys(recurrenceInfo.recurrence.rule[0].repeat[0]);
                            recurrenceTypes = [
                                "daily",
                                "weekly",
                                "monthly",
                                "monthlyByDay",
                                "yearly",
                                "yearlyByDay",
                            ];
                            freq = void 0;
                            for (_i = 0, keys_1 = keys; _i < keys_1.length; _i++) {
                                key = keys_1[_i];
                                rule = recurrenceInfo.recurrence.rule[0].repeat[0][key][0]["$"];
                                switch (recurrenceTypes.indexOf(key)) {
                                    case 0:
                                        // freq = RRule.DAILY;
                                        return [2 /*return*/, this.parseDailyRule(rule)];
                                        break;
                                    case 1:
                                        // freq = RRule.WEEKLY
                                        return [2 /*return*/, this.parseWeeklyRule(rule)];
                                        break;
                                    case 2:
                                        // freq = RRule.MONTHLY
                                        return [2 /*return*/, this.parseMonthlyRule(rule)];
                                        break;
                                    case 3:
                                        return [2 /*return*/, this.parseMonthlyByDayRule(rule)];
                                        break;
                                    case 4:
                                        // freq = RRule.WEEKLY
                                        return [2 /*return*/, this.parseYearlyRule(rule)];
                                        break;
                                    case 5:
                                        return [2 /*return*/, this.parseYearlyByDayRule(rule)];
                                        break;
                                    default:
                                        continue;
                                }
                            }
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
       *
       *
       * @private
       * @param {string} rule
       * @memberof Event
       */
    parseRecurrentEvent.prototype.parseDailyRule = function (rule) {
        var keys = Object.keys(rule);
        if (keys.indexOf("weekday") !== -1 && rule["weekday"] === "TRUE")
            return format("{} {}", format(strings.everyFormat, 1), strings.weekDayLabel);
        if (keys.indexOf("dayFrequency") !== -1) {
            var dayFrequency = parseInt(rule["dayFrequency"]);
            var frequencyFormat = dayFrequency === 1
                ? strings.everyFormat
                : dayFrequency === 2
                    ? strings.everySecondFormat
                    : strings.everyNthFormat;
            return format("{} {}", format(frequencyFormat, dayFrequency), strings.dayLable);
        }
        return "Invalid recurrence format";
    };
    /**
     *
     *
     * @private
     * @param { string } rule
     * @memberof Event
     */
    parseRecurrentEvent.prototype.parseWeeklyRule = function (rule) {
        var frequency = parseInt(rule["weekFrequency"]);
        var keys = Object.keys(rule);
        var dayMap = {
            mo: strings.Monday,
            tu: strings.Tuesday,
            we: strings.Wednesday,
            th: strings.Thursday,
            fr: strings.Friday,
            sa: strings.Saturday,
            su: strings.Sunday,
        };
        var days = [];
        for (var _i = 0, keys_2 = keys; _i < keys_2.length; _i++) {
            var key = keys_2[_i];
            days.push(dayMap[key]);
        }
        return format("{}{} {} {}", frequency === 1
            ? format(strings.everyFormat, frequency)
            : frequency === 2
                ? format(strings.everySecondFormat, frequency)
                : format(strings.everyNthFormat, frequency), strings.weekLabel, strings.onLabel, days.join(", "));
    };
    /**
     *
     *
     * @private
     * @param { string } rule
     * @memberof Event
     */
    parseRecurrentEvent.prototype.parseMonthlyRule = function (rule) {
        var frequency = parseInt(rule["monthFrequency"]);
        var day = parseInt(rule["day"]);
        return format("{}{} {}", frequency === 1
            ? format(strings.everyFormat, frequency)
            : frequency === 2
                ? format(strings.everySecondFormat, frequency)
                : format(strings.everyNthFormat, frequency), strings.monthLabel, format(strings.onTheDayFormat, day));
    };
    /**
     *
     * @private
     * @param { string } rule
     * @memberof Event
     */
    parseRecurrentEvent.prototype.parseMonthlyByDayRule = function (rule) {
        var keys = Object.keys(rule);
        var dayTypeMap = {
            day: strings.weekDayLabel,
            weekend_day: strings.weekEndDay,
            mo: strings.Monday,
            tu: strings.Tuesday,
            we: strings.Wednesday,
            th: strings.Thursday,
            fr: strings.Friday,
            sa: strings.Saturday,
            su: strings.Sunday,
        };
        var orderType = {
            first: strings.firstLabel,
            second: strings.secondLabel,
            third: strings.thirdLabel,
            fourth: strings.fourthLabel,
            last: strings.lastLabel,
        };
        var order;
        var dayType;
        var frequencyFormat;
        for (var _i = 0, keys_3 = keys; _i < keys_3.length; _i++) {
            var key = keys_3[_i];
            var frequency = parseInt(rule[key]);
            switch (key) {
                case "monthFrequency":
                    switch (frequency) {
                        case 1:
                            frequencyFormat = format(strings.everyFormat, frequency);
                            break;
                        case 2:
                            frequencyFormat = format(strings.everySecondFormat, frequency);
                            break;
                        default:
                            frequencyFormat = format(strings.everyNthFormat, frequency);
                            break;
                    }
                    break;
                case "weekDayOfMonth":
                    order = orderType[rule[key]];
                    break;
                default:
                    dayType = dayTypeMap[rule[key]];
                    break;
            }
        }
        return format("{} {} {} {} {}{}", frequencyFormat, strings.monthLabel.toLowerCase(), strings.onTheLabel, order, dayType, strings.theSuffix);
    };
    /**
     *
     * @private
     * @param rule
     * @memberof Event
     */
    parseRecurrentEvent.prototype.parseYearlyRule = function (rule) {
        var keys = Object.keys(rule);
        var months = DayPickerStrings.months;
        var frequencyString;
        var month;
        var day;
        for (var _i = 0, keys_4 = keys; _i < keys_4.length; _i++) {
            var key = keys_4[_i];
            var frequency = parseInt(rule[key]);
            var frequencyFormat = frequency === 1
                ? strings.everyFormat
                : frequency === 2
                    ? strings.everySecondFormat
                    : strings.everyNthFormat;
            switch (key) {
                case "yearFrequency":
                    frequencyString = format(frequencyFormat, frequency);
                    break;
                case "month":
                    month = months[parseInt(rule[key]) - 1];
                    break;
                case "day":
                    day = rule[key];
                    break;
            }
        }
        return format("{} {} {}", frequencyString, strings.yearLabel, format(strings.theNthOfMonthFormat, month, day));
    };
    /**
     *
     *
     * @private
     * @param rule
     * @memberof Event
     */
    parseRecurrentEvent.prototype.parseYearlyByDayRule = function (rule) {
        var keys = Object.keys(rule);
        var months = DayPickerStrings.months;
        var orderMap = {
            first: strings.firstLabel,
            second: strings.secondLabel,
            third: strings.thirdLabel,
            fourth: strings.fourthLabel,
            last: strings.lastLabel,
        };
        var dayTypeMap = {
            day: strings.weekDayLabel,
            weekend_day: strings.weekEndDay,
            mo: strings.Monday,
            tu: strings.Tuesday,
            we: strings.Wednesday,
            th: strings.Thursday,
            fr: strings.Friday,
            sa: strings.Saturday,
            su: strings.Sunday,
        };
        var frequencyString;
        var month;
        var order;
        var dayTypeString;
        for (var _i = 0, keys_5 = keys; _i < keys_5.length; _i++) {
            var key = keys_5[_i];
            var frequency = parseInt(rule[key]);
            var frequencyFormat = frequency === 1
                ? strings.everyFormat
                : frequency === 2
                    ? strings.everySecondFormat
                    : strings.everyNthFormat;
            switch (key) {
                case "yearFrequency":
                    frequencyString = format(frequencyFormat, frequency);
                    break;
                case "weekDayOfMonth":
                    order = orderMap[rule[key]];
                    break;
                case "month":
                    month = months[parseInt(rule[key]) - 1];
                    break;
                default:
                    dayTypeString = dayTypeMap[rule[key]];
                    break;
            }
            return format("{} {} {}", frequencyString, strings.yearLabel, format(strings.onTheDayTypeFormat, order, dayTypeString.toLowerCase(), strings.theSuffix));
        }
    };
    return parseRecurrentEvent;
}());
export default parseRecurrentEvent;
//# sourceMappingURL=parseRecurrentEvent.js.map