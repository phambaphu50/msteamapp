// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import path from 'path';
import restify from 'restify';
import { adapter, EchoBot } from './bot';
import tabs from './tabs';
import MessageExtension from './message-extension';

// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { ActivityTypes } from 'botbuilder';

// Read botFilePath and botFileSecret from .env file.
const ENV_FILE = path.join(__dirname, '.env');
require('dotenv').config({ path: ENV_FILE });

//Create HTTP server.
const server = restify.createServer({
    formatters: {
        'text/html': function (req, res, body) {
            return body;
        },
    },
});

server.get(
    '/*',
    restify.plugins.serveStatic({
        directory: __dirname + '/static',
    })
);

server.listen(process.env.port || process.env.PORT || 3333, function () {
    console.log(`\n${server.name} listening to ${server.url}`);
});

// Adding tabs to our app. This will setup routes to various views
tabs(server);

// Adding a bot to our app
const bot = new EchoBot();

// Adding a messaging extension to our app
const messageExtension = new MessageExtension();

// Listen for incoming requests.
server.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        if (context.activity.type === ActivityTypes.Invoke)
            await messageExtension.run(context);
        else await bot.run(context);
    });
});

// Listen for incoming notifications and send proactive messages to users.
server.get('/api/notify', async (req, res) => {
    for (const conversationReference of Object.values(conversationReferences)) {
        await adapter.continueConversation(conversationReference, async turnContext => {
            // If you encounter permission-related errors when sending this message, see
            // https://aka.ms/BotTrustServiceUrl
            await turnContext.sendActivity('proactive hello');
        });
    }

    res.setHeader('Content-Type', 'text/html');
    res.writeHead(200);
    res.write('<html><body><h1>Proactive messages have been sent.</h1></body></html>');
    res.end();
});