// Example configuration file
// Copy this to config.js and update with your Azure AD credentials

module.exports = {
  azure: {
    clientId: 'YOUR_CLIENT_ID_HERE', // Get this from Azure Portal
    authority: 'https://login.microsoftonline.com/common',
    scopes: ['User.Read', 'Calendars.Read']
  },
  widget: {
    width: 400,
    height: 600,
    position: 'top-right', // top-right, top-left, bottom-right, bottom-left
    autoStart: true
  }
};
