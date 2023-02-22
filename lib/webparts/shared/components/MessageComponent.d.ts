import * as React from 'react';
import { MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
/**
 * Interface to implement the MessageComponent Webpart
 *
 * @export
 * @interface IMessageComponentProps
 */
export interface IMessageComponentProps {
    Message: string;
    Type: MessageBarType;
    Display: boolean;
}
/**
 * React MessageComponent for displaying the messages
 *
 * @export
 * @class MessageComponent
 * @extends {React.Component<IMessageComponentProps, any>}
 */
export declare const MessageComponent: React.FunctionComponent<IMessageComponentProps>;
//# sourceMappingURL=MessageComponent.d.ts.map