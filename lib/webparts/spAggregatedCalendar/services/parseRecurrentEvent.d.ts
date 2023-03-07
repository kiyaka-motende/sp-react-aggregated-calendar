export default class parseRecurrentEvent {
    deCodeHtmlEntities(string: string): Promise<string>;
    /**
     *
     *
     * @private
     * @param {string} recurrenceData
     * @memberof Event
     */
    returnExceptionRecurrenceInfo(recurrenceData: string): Promise<string>;
    /**
       *
       *
       * @private
       * @param {string} rule
       * @memberof Event
       */
    parseDailyRule(rule: any): string;
    /**
     *
     *
     * @private
     * @param { string } rule
     * @memberof Event
     */
    parseWeeklyRule(rule: any): string;
    /**
     *
     *
     * @private
     * @param { string } rule
     * @memberof Event
     */
    parseMonthlyRule(rule: {
        [x: string]: string;
    }): string;
    /**
     *
     * @private
     * @param { string } rule
     * @memberof Event
     */
    parseMonthlyByDayRule(rule: any): string;
    /**
     *
     * @private
     * @param rule
     * @memberof Event
     */
    parseYearlyRule(rule: any): string;
    /**
     *
     *
     * @private
     * @param rule
     * @memberof Event
     */
    parseYearlyByDayRule(rule: any): string;
}
//# sourceMappingURL=parseRecurrentEvent.d.ts.map