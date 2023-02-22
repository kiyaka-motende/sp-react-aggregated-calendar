import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
/**
 * React MessageComponent for displaying the messages
 *
 * @export
 * @class MessageComponent
 * @extends {React.Component<IMessageComponentProps, any>}
 */
export var MessageComponent = function (props) {
    /**
    * Render method of the Message Component
    *
    * @returns {React.ReactElement<IMessageComponentProps>}
    * @memberof MessageComponent
    */
    return (React.createElement("div", { className: "ms-Grid-row" },
        React.createElement("div", { className: "ms-Grid-col ms-sm12" }, props.Display &&
            React.createElement("div", null,
                React.createElement(MessageBar, { messageBarType: MessageBarType.error, isMultiline: false, dismissButtonAriaLabel: "Close" }, props.Message)))));
};
//# sourceMappingURL=MessageComponent.js.map