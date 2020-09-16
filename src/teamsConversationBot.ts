// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import {
    ActionTypes,
    BotFrameworkAdapter,
    CardFactory,
    ChannelAccount,
    MessageFactory,
    TeamInfo,
    TeamsActivityHandler,
    TeamsInfo,
    TurnContext
} from 'botbuilder';
const TextEncoder = require( 'util' ).TextEncoder;

export class TeamsConversationBot extends TeamsActivityHandler {
    constructor() {
        super();
        this.onMessage( async ( context: TurnContext, next ): Promise<void> => {
            TurnContext.removeRecipientMention( context.activity );
            const text = context.activity.text.trim().toLocaleLowerCase();
            if ( text.includes( 'draw names' ) ) {
                await this.sendDrawNamesCard(context);
            } 
            await next();
        } );
    }


    public async sendDrawNamesCard( context: TurnContext): Promise<void> {

        const teamMembers = await TeamsInfo.getMembers(context);

        const randomizedNames = randomize(teamMembers.map(x => x.name.toString()));

        const renderedNames = `<strong><big><ol>${ randomizedNames.map(x => `<li>${x}</li>`).join('') }</ol></big></strong>`

        const card = CardFactory.heroCard(
            'The Goblet of Fire has spoken!',
            renderedNames,
            ['https://media.giphy.com/media/UsVK0H1hSRi1y/giphy.gif'],
            null
        );
        await context.sendActivity( MessageFactory.attachment( card ) );
    }
}

/**
 * Randomize the order of a list
 * @param list 
 */
function randomize<T>(list: T[]): T[] {
    return [...list].sort((a, b) => (Math.random() > 0.5 ? 1 : -1));
}