export interface IEventData {
    Id?: number;
    ID?: number;
    title: string;
    Description?: any;
    location?: string;
    start: Date;
    end: Date;
    color?: string;
    ownerInitial?: string;
    ownerPhoto?: string;
    ownerEmail?: string;
    ownerName?: string;
    fAllDayEvent?: boolean;
    attendes?: number[];
    geolocation?: {
        Longitude: number;
        Latitude: number;
    };
    Category?: string;
    Duration?: number;
    RecurrenceData?: string;
    rrule?: any;
    fRecurrence?: string | boolean;
    EventType?: string;
    UID?: string;
    RecurrenceID?: Date;
    MasterSeriesItemID?: string;
}
//# sourceMappingURL=IEventData.d.ts.map