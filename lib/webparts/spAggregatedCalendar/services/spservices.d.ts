import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IEventData } from "./IEventData";
export default class spservices {
    private context;
    constructor(context: WebPartContext);
    private onInit;
    /**
     *
     * @private
     * @returns {Promise<string>}
     * @memberof spservices
     */
    getLocalTime(date: string | Date): Promise<string>;
    /**
     *
     * @private
     * @returns {Promise<string>}
     * @memberof spservices
     */
    getUtcTime(date: string | Date): Promise<string>;
    /**
     *
     * @param {string} siteUrl
     * @param {string} listId
     * @param {string} fieldInternalName
     * @returns {Promise<{ key: string, text: string }[]>}
     * @memberof spservices
     */
    getChoiceFieldOptions(siteUrl: string, listId: string, fieldInternalName: string): Promise<{
        key: string;
        text: string;
    }[]>;
    /**
     *
     * @param {string} siteUrl
     * @param {string} listId
     * @param {Date} eventStartDate
     * @param {Date} eventEndDate
     * @returns {Promise< IEventData[]>}
     * @memberof spservices
     */
    getEvents(siteUrl: string, listTitle: string, eventStartDate: Date, eventEndDate: Date): Promise<IEventData[]>;
}
//# sourceMappingURL=spservices.d.ts.map