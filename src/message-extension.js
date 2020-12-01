import { TeamsActivityHandler, CardFactory } from 'botbuilder';
import faker from 'faker';

export default class MessageExtension extends TeamsActivityHandler {
    async handleTeamsMessagingExtensionQuery(context, query) {
        // If the user supplied a title via the cardTitle parameter then use it or use a fake title
        let title =
            query.parameters && query.parameters[0].name === 'cardTitle'
                ? query.parameters[0].value
                : faker.lorem.sentence();

        let randomImageUrl = 'https://loremflickr.com/200/200'; // Faker's random images uses lorempixel.com, which has been down a lot

        switch (query.commandId) {
            case 'getRandomText':
                let attachments = [];

                // Generate 5 results to send with fake text and fake images
                for (let i = 0; i < 5; i++) {
                    let text = faker.lorem.paragraph();
                    let images = [`${randomImageUrl}?random=${i}`];
                    let thumbnailCard = CardFactory.thumbnailCard(
                        title,
                        text,
                        images
                    );
                    let preview = CardFactory.thumbnailCard(
                        title,
                        text,
                        images
                    );
                    preview.content.tap = {
                        type: 'invoke',
                        value: { title, text, images },
                    };
                    var attachment = { ...thumbnailCard, preview };
                    attachments.push(attachment);
                }

                return {
                    composeExtension: {
                        type: 'result',
                        attachmentLayout: 'list',
                        attachments: attachments,
                    },
                };
            case 'getRandomQuestion':
                let cards = [];

                for (let i = 0; i < 5; i++) {
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
                                text: ' ${location} ',
                                isSubtle: true,
                                wrap: true,
                            },
                            {
                                type: 'TextBlock',
                                text:
                                    "${formatDateTime(start.dateTime, 'HH:mm')} - ${formatDateTime(end.dateTime, 'hh:mm')}",
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
                                value: '${reminders.overrides[0].minutes}',
                                choices: [
                                    {
                                        $data: '${reminders.overrides}',
                                        title: '${minutes} minutes',
                                        value: '${minutes}',
                                    },
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
                    cards.push(activeCard);
                }
                return {
                    composeExtension: {
                        type: 'result',
                        attachmentLayout: 'list',
                        attachments: cards,
                    },
                };
            default:
                break;
        }
        return null;
    }

    async handleTeamsMessagingExtensionSelectItem(context, obj) {
        const { title, text, images } = obj;

        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: [CardFactory.thumbnailCard(title, text, images)],
            },
        };
    }
}
