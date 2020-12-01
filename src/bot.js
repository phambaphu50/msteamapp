// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import {
    TeamsActivityHandler,
    BotFrameworkAdapter,
    MemoryStorage,
    ConversationState,
    TurnContext,
    CardFactory,
} from 'botbuilder';
import config from 'config';

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
export const adapter = new BotFrameworkAdapter({
    appId: config.get('bot.appId'),
    appPassword: config.get('bot.appPassword'),
});

adapter.onTurnError = async (context, error) => {
    const errorMsg = error.message
        ? error.message
        : `Oops. Something went wrong!`;
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError] unhandled error: ${error}`);

    // Clear out state
    await conversationState.delete(context);
    // Send a message to the user
    await context.sendActivity(errorMsg);

    // Note: Since this Messaging Extension does not have the messageTeamMembers permission
    // in the manifest, the bot will not be allowed to message users.
};

// Define state store for your bot.
const memoryStorage = new MemoryStorage();

// Create conversation state with in-memory storage provider.
const conversationState = new ConversationState(memoryStorage);

export class EchoBot extends TeamsActivityHandler {
    constructor() {
        super();
        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);
            const text = context.activity.text.trim().toLocaleLowerCase();
            // await context.sendActivity('You said ' + text);
            if (text.includes('questions')) {
                await this.runQuestion(context);
            }
        });
    }
    async runQuestion(context) {
        const card = {
            type: 'AdaptiveCard',
            body: [
                {
                    type: 'TextBlock',
                    text: 'summary',
                    size: 'Large',
                    weight: 'Bolder',
                    wrap: true,
                },
                {
                    type: 'TextBlock',
                    text: '',
                    isSubtle: true,
                    wrap: true,
                },
                {
                    type: 'TextBlock',
                    text:
                        "",
                    isSubtle: true,
                    spacing: 'None',
                    wrap: true,
                },
                {
                    type: 'TextBlock',
                    text: 'Snooze for',
                    wrap: true,
                },
                {
                    type: 'Input.ChoiceSet',
                    id: 'snooze',
                    value: '1',
                    choices: [
                        {
                            title: '1',
                            value: '1',
                        },
                        {
                            title: '2',
                            value: '2',
                        }
                    ],
                },
            ],
            actions: [
                {
                    type: 'Action.Submit',
                    title: 'Snooze',
                    data: {
                        x: 'snooze',
                    },
                },
                {
                    type: 'Action.Submit',
                    title: "I'll be late",
                    data: {
                        x: 'late',
                    },
                },
            ],
            version: '1.0.0',
        };
        const activeCard = CardFactory.adaptiveCard(card);
        await context.sendActivity({
            text: 'Please give a minutes to answer the questions!', 
            attachments: [activeCard]
        });
    }
}
