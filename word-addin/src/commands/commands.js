/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
    // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function showTaskpane(event) {
    // The following example shows task pane functionality.
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "TabHome",
                groups: [
                    {
                        id: "CommandsGroup",
                        controls: [
                            {
                                id: "TaskpaneButton",
                                enabled: true
                            }
                        ]
                    }
                ]
            }
        ]
    });

    event.completed();
}

// Register the function with Office
Office.actions.associate("showTaskpane", showTaskpane); 