require('isomorphic-fetch');
const azure = require('@azure/identity');
const graph = require('@microsoft/microsoft-graph-client');
const authProviders =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

let _settings = undefined;
let _deviceCodeCredential = undefined;
let _userClient = undefined;

const body_content = `<div id="email" style="width:600px;margin: auto;background:white;">
<table role="presentation" border="0" align="right" cellspacing="0">
<tr>
    <td>
    <a href="#" style="font-size: 9px; text-transform:uppercase; letter-spacing: 1px; color: #99ACC2;  font-family:Avenir;">View in Browser</a>
    </td>
</tr>
</table>

<!-- Header --> 
<table role="presentation" border="1" width="100%" cellspacing="0">
<tr>
<td bgcolor="white" align="center" style="color: #00A4BD;">
    <img alt="Flower" src="https://i.ibb.co/6XZwsDL/liv-email-01.png" width="80%" style="padding-top : 20px">
    <h1 style="font-size: 52px; margin:0 0 20px 0; font-family:Avenir;"></h1>
</tr>
    </td>
</table>

<!-- Body 1 --> 
<table role="presentation"  bgColor="white" border="1" width="100%" cellspacing="0">
<tr>
    <td style="padding: 30px 30px 30px 60px;">
    <h2 style="font-size: 28px; margin:0 0 20px 0; font-family:Avenir;">Bay Area Orders</h2>
    <p style="margin:0 0 12px 0;font-size:16px;line-height:24px;font-family:Avenir">Hi None,



        Hope you are doing well! I wanted to connect with you on the current inventory status for your region. The ‘Need to Order’ quantities below are based on the number of warehouses that have or will dip below our MDT of 300 units (1 pallet) by 2022-11-20.</p>
    </td> 
</tr>
</table>

<!-- Body 2--> 
<table role="presentation" border="1" width="100%" cellspacing="0" >
    <tr>
        <td style="vertical-align: top; padding: 10px; display: flex; justify-content: center;"> 
        <table style="width: 90%; border-collapse: collapse; border-radius:6px; margin: 25px 0; font-size: 0.9em; font-family: sans-serif; min-width: 400px; box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);">
        <thead>
            <tr style="background-color: #0b5da3; color: #ffffff; text-align: left;">
            <th style="border-top-left-radius: 6px;padding: 12px 15px;">DC</th>
            <th>Product</th>
            <th>MDT</th>
            <th>NTO</th>
            <th style="border-top-right-radius: 6px;padding: 12px 15px;">Week</th>
            </tr>
        </thead>
<tbody style="border-bottom: 2px orange"><tr style="border-bottom: 1px solid #dddddd;"><td>1354 KATY DRY</td><td style="background-color: #7fff78; padding: 12px 15px">Energy</td><td style="padding: 12px 15px">1</td><td style="padding: 12px 15px">0</td><td style="padding: 12px 15px">2022-11-20</td></tr><tr style="background-color: #f3f3f3; border-bottom: 1px solid #dddddd;"><td>1354 KATY DRY</td><td style="background-color: #87fff9; padding: 12px 15px">Immune Support</td><td style="padding: 12px 15px">4</td><td style="padding: 12px 15px">0</td><td style="padding: 12px 15px">2022-11-20</td></tr><tr style="border-bottom: 1px solid #dddddd;"><td>1354 KATY DRY</td><td style="background-color: #fcef77; padding: 12px 15px">Lemon Lime</td><td style="padding: 12px 15px">1</td><td style="padding: 12px 15px">0</td><td style="padding: 12px 15px">2022-11-20</td></tr><tr style="background-color: #f3f3f3; border-bottom: 1px solid #dddddd;"><td>1354 KATY DRY</td><td style="background-color: #ff7f78; padding: 12px 15px">Strawberry</td><td style="padding: 12px 15px">1</td><td style="padding: 12px 15px">0</td><td style="padding: 12px 15px">2022-11-20</td></tr><tr style="border-bottom: 1px solid #dddddd;"><td>288 DALLAS DRY</td><td style="background-color: #7fff78; padding: 12px 15px">Energy</td><td style="padding: 12px 15px">1</td><td style="padding: 12px 15px">0</td><td style="padding: 12px 15px">2022-11-20</td></tr><tr style="background-color: #f3f3f3; border-bottom: 1px solid #dddddd;"><td>288 DALLAS DRY</td><td style="background-color: #87fff9; padding: 12px 15px">Immune Support</td><td style="padding: 12px 15px">1</td><td style="padding: 12px 15px">0</td><td style="padding: 12px 15px">2022-11-20</td></tr><tr style="border-bottom: 1px solid #dddddd;"><td>288 DALLAS DRY</td><td style="background-color: #fcef77; padding: 12px 15px">Lemon Lime</td><td style="padding: 12px 15px">1</td><td style="background-color: orange; padding: 12px 15px">1</td><td style="padding: 12px 15px">2022-11-20</td></tr><tr style="background-color: #f3f3f3; border-bottom: 1px solid #dddddd;border-bottom: 2px solid #1e6098"><td>288 DALLAS DRY</td><td style="background-color: #ff7f78; padding: 12px 15px">Strawberry</td><td style="padding: 12px 15px">1</td><td style="padding: 12px 15px">0</td><td style="padding: 12px 15px">2022-11-20</td></tr></tbody>
</table>
</td>
</tr>       
</table>

<!-- Body 3 --> 
<table role="presentation" border="1" width="100%">
<tr>
<td bgcolor="#EAF0F6" align="center" style="padding: 30px 30px;">
<h2 style="font-size: 28px; margin:0 0 20px 0; font-family:Avenir;"> Attached Below </h2>
<p style="margin:0 0 12px 0;font-size:16px;line-height:24px;font-family:Avenir">I have attached the full list of warehouses with their inventory tracking for your reference. Please let me know if you are aligned and if we should be expecting any new PO’s. Thanks so much!</p>
</td>
</tr>
</table>

<!-- Footer -->
<table role="presentation" border="1" width="100%" cellspacing="0">
<tr>
<td bgcolor="#F5F8FA" style="padding: 30px 30px;">
    <p style="margin:0 0 12px 0; font-size:16px; line-height:24px; color: #99ACC2; font-family:Avenir"> Liquid I.V. </p>
    <p style="font-size: 9px; text-transform:uppercase; letter-spacing: 1px; color: #99ACC2;  font-family:Avenir;"> Fueling Life's Adventures </p>      
</td>
</tr>
</table> 
</div>`

function initializeGraphForUserAuth(settings, deviceCodePrompt) {
  // Ensure settings isn't null
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }

  _settings = settings;

  _deviceCodeCredential = new azure.DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.authTenant,
    userPromptCallback: deviceCodePrompt
  });

  const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
    _deviceCodeCredential, {
      scopes: settings.graphUserScopes
    });

  _userClient = graph.Client.initWithMiddleware({
    authProvider: authProvider
  });
}
async function getUserTokenAsync() {
    // Ensure credential isn't undefined
    if (!_deviceCodeCredential) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    // Ensure scopes isn't undefined
    if (!_settings?.graphUserScopes) {
      throw new Error('Setting "scopes" cannot be undefined');
    }
  
    // Request token with given scopes
    const response = await _deviceCodeCredential.getToken(_settings?.graphUserScopes);
    return response.token;
  }


  async function getUserTokenAsync() {
    // Ensure credential isn't undefined
    if (!_deviceCodeCredential) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    // Ensure scopes isn't undefined
    if (!_settings?.graphUserScopes) {
      throw new Error('Setting "scopes" cannot be undefined');
    }
  
    // Request token with given scopes
    const response = await _deviceCodeCredential.getToken(_settings?.graphUserScopes);
    return response.token;
  }

  async function getUserAsync() {
    // Ensure client isn't undefined
    if (!_userClient) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    return _userClient.api('/me')
      // Only request specific properties
      .select(['displayName', 'mail', 'userPrincipalName'])
      .get();
  }

  async function getInboxAsync() {
    // Ensure client isn't undefined
    if (!_userClient) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    return _userClient.api('/me/mailFolders/inbox/messages')
      .select(['from', 'isRead', 'receivedDateTime', 'subject'])
      .top(25)
      .orderby('receivedDateTime DESC')
      .get();
  }

  async function sendMailAsync(subject, body, recipient) {
    // Ensure client isn't undefined
    if (!_userClient) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    // Create a new message
    const message = {
      subject: subject,
      body: {
        content: body_content,
        contentType: 'html'
      },
      toRecipients: [
        {
          emailAddress: {
            address: recipient
          }
        }
      ]
    };
  
    // Send the message
    return _userClient.api('me/sendMail')
      .post({
        message: message
      });
  }

module.exports.sendMailAsync = sendMailAsync;
module.exports.getInboxAsync = getInboxAsync;
module.exports.getUserAsync = getUserAsync;
module.exports.getUserTokenAsync = getUserTokenAsync;
module.exports.getUserTokenAsync = getUserTokenAsync;
module.exports.initializeGraphForUserAuth = initializeGraphForUserAuth;