// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TurnContext,
    MessageFactory,
    TeamsActivityHandler,
    CardFactory,
    ActionTypes,
    ConsoleTranscriptLogger
} = require('botbuilder');

class BotActivityHandler extends TeamsActivityHandler {
    constructor() {
        super();
        /*  Teams bots are Microsoft Bot Framework bots.
            If a bot receives a message activity, the turn handler sees that incoming activity
            and sends it to the onMessage activity handler.
            Learn more: https://aka.ms/teams-bot-basics.

            NOTE:   Ensure the bot endpoint that services incoming conversational bot queries is
                    registered with Bot Framework.
                    Learn more: https://aka.ms/teams-register-bot. 
        */
        // Registers an activity event handler for the message event, emitted for every incoming message activity.
        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);
            switch (context.activity.text.toLowerCase().trim()) {
            case 'hello':
                await this.mentionActivityAsync(context);
                break;
            case 'holidays':
                var text = getHolidays();
                await this.sendMessage(context, text);
                break;
            case ('hr portal'):
                var text = "https://medtronicprod.service-now.com/hr";
                await this.sendMessage(context, text);
                break;
            case ('recognize'):
                var text = "https://recognize.medtronic.com";
                await this.sendMessage(context, text);
                break;   
            case ('azure'):
                var text = "Azure DevOps: https://dev.azure.com \n\n Azure Portal: https://portal.azure.com";
                await this.sendMessage(context, text);
                break;
            case ('benefits'):
                var text = "Medtronic Benefits Page: https://www3.benefitsolver.com/benefits/BenefitSolverView \n\n"
                            + "Fidelity NetBenefits: https://nb.fidelity.com/public/nb/default/home \n\n"
                            + "MEAP: https://medtronicprod.service-now.com/hr?id=subpage&sysparm_id=MEAP \n\n"
                            + "Healthier Together: https://landing.virginpulse.com/Medtronic/ \n\n"
                            + "Privacy Armor: https://privacyarmor.infoarmor.com/signin/?partnerid=infoarmor";
                await this.sendMessage(context, text);
                break;
            case ('resources'):
                var text = "Plural Site: https://app.pluralsight.com/id/ \n\n"
                            + "Acronym Finder: https://medtronic.sharepoint.com/sites/knowledgecenter/sitepages/acronymfinder.aspx# \n\n"
                            + "Your Cause: https://medtronic.yourcause.com/#/home";
                await this.sendMessage(context, text);
                break;   
            default:
                // By default for unknown activity sent by user show
                // a card with the available actions.
                const value = { count: 0 };
                const card = CardFactory.heroCard(
                    'Lets talk...',
                    null,
                    [{
                        type: ActionTypes.MessageBack,
                        title: 'Say Hello',
                        value: value,
                        text: 'Hello'
                    }]);
                await context.sendActivity({ attachments: [card] });
                break;
            }
            await next();
        });
    }

    /**
     * Say hello and @ mention the current user.
     */
    async mentionActivityAsync(context) {
        const TextEncoder = require('html-entities').XmlEntities;

        const mention = {
            mentioned: context.activity.from,
            text: `<at>${ new TextEncoder().encode(context.activity.from.name) }</at>`,
            type: 'mention'
        };

        const replyActivity = MessageFactory.text(`Hi ${ mention.text }`);
        replyActivity.entities = [mention];
        
        await context.sendActivity(replyActivity);
    }

    async sendMessage(context, text) {
        const TextEncoder = require('html-entities').XmlEntities;

        const mention = {
            mentioned: context.activity.from,
            text: text,
            type: 'mention'
        };

        const replyActivity = MessageFactory.text(`${ mention.text }`);
        replyActivity.entities = [mention];
        await context.sendActivity(replyActivity);
    }
}



module.exports.BotActivityHandler = BotActivityHandler;

function getHolidays(){
    holidays = ["Friday, January 1: New Year's Day",
                "Monday, January 18: Martin Luther King Jr. Day",
                "Monday, May 31: Memorial Day",
                "Monday, July 5: Independence Day (observed)",
                "Monday, September 6: Labor Day",
                "Thursday, November 25: Thanksgiving Day",
                "Friday, November 26: Day After Thanksgiving",
                "Friday, December 24: Christmas Eve",
                "Monday, December 27: Christmas Day (observed)",
                "Friday, December 31: New Year's Eve"];
    return holidays.join('\n\n');
}
