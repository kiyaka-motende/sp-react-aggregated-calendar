/* eslint-disable dot-notation */
/* eslint-disable @typescript-eslint/no-var-requires */
// import { RRule } from "rrule";
import { parseString } from "xml2js";
import * as strings from "SpAggregatedCalendarWebPartStrings";
import { IDatePickerStrings } from "@fluentui/react";
const format: any = require("string-format");

const DayPickerStrings: IDatePickerStrings = {
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

export default class parseRecurrentEvent {
    public async deCodeHtmlEntities(string: string) {

        const HtmlEntitiesMap = {
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
    
        const entityMap:any = HtmlEntitiesMap;
        for (const key in entityMap) {
            if (Object.prototype.hasOwnProperty.call(entityMap, key)) {
                const entity = entityMap[key];
                const regex = new RegExp(entity, 'g');
                string = string.replace(regex, key);
            }
        }
        string = string.replace(/&quot;/g, '"');
        string = string.replace(/&amp;/g, '&');
        return string;
      }

  /**
   *
   *
   * @private
   * @param {string} recurrenceData
   * @memberof Event
   */
   public async returnExceptionRecurrenceInfo(recurrenceData: string) {
    const promise = new Promise<object>((resolve, reject) => {
      parseString(recurrenceData, (err, result) => {
        if (err) {
          reject(err);
        }

        resolve(result);
      });
    });

    const recurrenceInfo: any = await promise;
    if (recurrenceInfo !== null) {
      const keys = Object.keys(recurrenceInfo.recurrence.rule[0].repeat[0]);
      const recurrenceTypes = [
        "daily",
        "weekly",
        "monthly",
        "monthlyByDay",
        "yearly",
        "yearlyByDay",
      ];

      let freq
      for (const key of keys) {
        const rule = recurrenceInfo.recurrence.rule[0].repeat[0][key][0]["$"];
        switch (recurrenceTypes.indexOf(key)) {
          case 0:
            // freq = RRule.DAILY;
            return this.parseDailyRule(rule);
            break;
          case 1:
            // freq = RRule.WEEKLY
            return this.parseWeeklyRule(rule);
            break;
          case 2:
            // freq = RRule.MONTHLY
            return this.parseMonthlyRule(rule);
            break;
          case 3:
            return this.parseMonthlyByDayRule(rule);
            break;
          case 4:
            // freq = RRule.WEEKLY
            return this.parseYearlyRule(rule);
            break;
          case 5:
            return this.parseYearlyByDayRule(rule);
            break;
          default:
            continue;
        }
      }
    }
  }

/**
   *
   *
   * @private
   * @param {string} rule
   * @memberof Event
   */
 public parseDailyRule(rule:any): string {
    const keys = Object.keys(rule);
    if (keys.indexOf("weekday") !== -1 && rule["weekday"] === "TRUE")
      return format(
        "{} {}",
        format(strings.everyFormat, 1),
        strings.weekDayLabel
      );

    if (keys.indexOf("dayFrequency") !== -1) {
      const dayFrequency: number = parseInt(rule["dayFrequency"]);
      const frequencyFormat =
        dayFrequency === 1
          ? strings.everyFormat
          : dayFrequency === 2
          ? strings.everySecondFormat
          : strings.everyNthFormat;
      return format(
        "{} {}",
        format(frequencyFormat, dayFrequency),
        strings.dayLable
      );
    }

    return "Invalid recurrence format";
  }

  /**
   *
   *
   * @private
   * @param { string } rule
   * @memberof Event
   */
  public parseWeeklyRule(rule: any): string {
    const frequency: number = parseInt(rule["weekFrequency"]);
    const keys = Object.keys(rule);
    const dayMap: any = {
      mo: strings.Monday,
      tu: strings.Tuesday,
      we: strings.Wednesday,
      th: strings.Thursday,
      fr: strings.Friday,
      sa: strings.Saturday,
      su: strings.Sunday,
    };
    const days: string[] = [];
    for (const key of keys) {
      days.push(dayMap[key]);
    }

    return format(
      "{}{} {} {}",
      frequency === 1
        ? format(strings.everyFormat, frequency)
        : frequency === 2
        ? format(strings.everySecondFormat, frequency)
        : format(strings.everyNthFormat, frequency),
      strings.weekLabel,
      strings.onLabel,
      days.join(", ")
    );
  }

  /**
   *
   *
   * @private
   * @param { string } rule
   * @memberof Event
   */
  public parseMonthlyRule(rule: { [x: string]: string; }): string {
    const frequency: number = parseInt(rule["monthFrequency"]);
    const day: number = parseInt(rule["day"]);

    return format(
      "{}{} {}",
      frequency === 1
        ? format(strings.everyFormat, frequency)
        : frequency === 2
        ? format(strings.everySecondFormat, frequency)
        : format(strings.everyNthFormat, frequency),
      strings.monthLabel,
      format(strings.onTheDayFormat, day)
    );
  }

  /**
   *
   * @private
   * @param { string } rule
   * @memberof Event
   */
  public parseMonthlyByDayRule(rule:any): string {
    const keys: string[] = Object.keys(rule);
    const dayTypeMap: any = {
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

    const orderType: any = {
      first: strings.firstLabel,
      second: strings.secondLabel,
      third: strings.thirdLabel,
      fourth: strings.fourthLabel,
      last: strings.lastLabel,
    };

    let order: string;
    let dayType: string;
    let frequencyFormat: string;

    for (const key of keys) {
    const frequency = parseInt(rule[key]);
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

    return format(
      "{} {} {} {} {}{}",
      frequencyFormat,
      strings.monthLabel.toLowerCase(),
      strings.onTheLabel,
      order,
      dayType,
      strings.theSuffix
    );
  }

  /**
   *
   * @private
   * @param rule
   * @memberof Event
   */
  public parseYearlyRule(rule:any): string {
    const keys: string[] = Object.keys(rule);
    const months: string[] = DayPickerStrings.months;
    let frequencyString: string;
    let month: string;
    let day: string;
    for (const key of keys) {
        const frequency = parseInt(rule[key]);
        const frequencyFormat =
          frequency === 1
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

    return format(
      "{} {} {}",
      frequencyString,
      strings.yearLabel,
      format(strings.theNthOfMonthFormat, month, day)
    );
  }

  /**
   *
   *
   * @private
   * @param rule
   * @memberof Event
   */
  public parseYearlyByDayRule(rule:any): string {
    const keys: string[] = Object.keys(rule);
    const months: string[] = DayPickerStrings.months;
    const orderMap: any = {
      first: strings.firstLabel,
      second: strings.secondLabel,
      third: strings.thirdLabel,
      fourth: strings.fourthLabel,
      last: strings.lastLabel,
    };
    const dayTypeMap: any = {
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
    let frequencyString: string;
    let month: string;
    let order: string;
    let dayTypeString: string;
    for (const key of keys) {
        const frequency = parseInt(rule[key]);
        const frequencyFormat =
          frequency === 1
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

      return format(
        "{} {} {}",
        frequencyString,
        strings.yearLabel,
        format(
          strings.onTheDayTypeFormat,
          order,
          dayTypeString.toLowerCase(),
          strings.theSuffix
        )
      );
    }
  }

}