const { PublicClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

// You need to register an app in Azure AD and get these values
const msalConfig = {
  auth: {
    clientId: 'YOUR_CLIENT_ID', // Replace with your Azure AD app client ID
    authority: 'https://login.microsoftonline.com/common',
  }
};

const pca = new PublicClientApplication(msalConfig);

const scopes = ['User.Read', 'Calendars.Read'];

let cachedAccount = null;

async function getAuthClient() {
  try {
    const authResult = await pca.acquireTokenInteractive({
      scopes: scopes,
      prompt: 'select_account'
    });

    cachedAccount = authResult.account;
    return authResult.account;
  } catch (error) {
    console.error('Authentication error:', error);
    throw error;
  }
}

async function getAccessToken() {
  if (!cachedAccount) {
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length === 0) {
      throw new Error('No cached account found. Please login again.');
    }
    cachedAccount = accounts[0];
  }

  try {
    const response = await pca.acquireTokenSilent({
      scopes: scopes,
      account: cachedAccount
    });
    return response.accessToken;
  } catch (error) {
    console.error('Token acquisition error:', error);
    const response = await pca.acquireTokenInteractive({
      scopes: scopes,
      account: cachedAccount
    });
    return response.accessToken;
  }
}

async function getCalendarEvents(startDate, endDate) {
  try {
    const accessToken = await getAccessToken();

    const client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    const startDateTime = startDate.toISOString();
    const endDateTime = endDate.toISOString();

    const events = await client
      .api('/me/calendarview')
      .query({
        startDateTime: startDateTime,
        endDateTime: endDateTime,
        $orderby: 'start/dateTime'
      })
      .select('subject,start,end,location,onlineMeeting,organizer,body,isAllDay')
      .top(50)
      .get();

    return events.value.map(event => ({
      id: event.id,
      subject: event.subject,
      start: event.start.dateTime,
      end: event.end.dateTime,
      location: event.location?.displayName || '',
      teamsLink: event.onlineMeeting?.joinUrl || null,
      organizer: event.organizer?.emailAddress?.name || 'Unknown',
      body: event.body?.content || '',
      isAllDay: event.isAllDay || false
    }));
  } catch (error) {
    console.error('Error fetching calendar events:', error);
    throw error;
  }
}

module.exports = {
  getAuthClient,
  getCalendarEvents
};
