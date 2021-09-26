import {
  BotFrameworkAdapter
} from 'botbuilder';
import { MicrosoftAppCredentials } from 'botframework-connector';
import { ChannelAccount } from 'botframework-connector/lib/connectorApi/models';

const serviceUrl = '';
const botAppId = '';
const botAppPassword = '';

const teamId = '';

const adapter = new BotFrameworkAdapter({
  appId: botAppId,
  appPassword: botAppPassword
});


/**
   * Fetches all members for a given team id
   * @param teamId teamId
   */
 async function getAllTeamMembers(): Promise<ChannelAccount[]> {
  const client = adapter.createConnectorClient(serviceUrl);
  MicrosoftAppCredentials.trustServiceUrl(serviceUrl);
  const response = await client.conversations.getConversationMembers(teamId)
  .catch((error) => {
    console.error(`getAllTeamMembers - error occurred: ${JSON.stringify(error)}`);
    throw error;
  });

  if (response) {
    console.debug(`getAllTeamMembers - response = ${JSON.stringify(response)}`);
    return response as ChannelAccount[];
  } else {
    console.info(`getAllTeamMembers - No users were found for teamId: ${teamId}`);
    throw new Error('No users were found for team');
  }
}


getAllTeamMembers();