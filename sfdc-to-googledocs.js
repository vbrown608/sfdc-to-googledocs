/* Salesforce integration to Google Docs using OAuth2
 * Vivian Brown
 * 1-31-2013
 */

// Hardcoded variables. Will move this to ScriptProperties.
var AUTHORIZE_URL = 'https://login.salesforce.com/services/oauth2/authorize'; //Get authorization code here
var TOKEN_URL = 'https://login.salesforce.com/services/oauth2/token'; //Get access token here

// App-specific settings.
var CLIENT_ID = ''; // INSERT STRING CLIENT ID HERE
var CLIENT_SECRET= ''; // INSERT STRING CLIENT SECRET HERE
var REDIRECT_URL= ScriptApp.getService().getUrl();

// These are the User Properties where we'll store the token. 
// Make sure this is unique across all user properties across all scripts.
var baseURLPropertyName = 'SALESFORCE_INSTANCE_URL'; 
var accessTokenPropertyName = 'SALESFORCE_OAUTH_TOKEN'; 
var refreshTokenPropertyName = 'SALESFORCE_REFRESH_TOKEN';

// Create a drop-down menu in the spreadsheet application.
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Download from SalesForce", functionName: "salesforceEntryPoint"}];
  ss.addMenu("Salesforce.com", menuEntries);
}

// This function is called from the drop-down menu. It initiates the login to salesforce
// or request for data.
function salesforceEntryPoint(){
  var token = UserProperties.getProperty(accessTokenPropertyName)
  if (!token) { // No saved token - we're starting from scratch
    login();
  } else { // Try a data request and deal with any errors
    requestAndHandleData();
  }
}

// Log in to Salesforce and get a brand new access and refresh token.
function login() {
  var URLForAuthorization = AUTHORIZE_URL + '?response_type=code&client_id='+CLIENT_ID+'&redirect_uri='+REDIRECT_URL;
  var HTMLToOutput = "<html><h1>You need to login</h1><a href='"+URLForAuthorization+"'>Click here to authorize this app.</a><br>Re-open this window when you return.</html>";
  SpreadsheetApp.getActiveSpreadsheet().show(HtmlService.createHtmlOutput(HTMLToOutput));
}

// Callback - this is called by Salesforce after the user approves the app.
function doGet(e) {
  var HTMLToOutput;
  if(e.parameters.code){ //If we get the authorization code as a parameter in, then this is a callback. 
    getAndStoreAccessToken(e.parameters.code);
    HTMLToOutput = '<html><h1>Finished with oAuth</h1>You can close this window.</html>';
  }
  return HtmlService.createHtmlOutput(HTMLToOutput);
}

// Use the authorization code the get the instance url and access token
function getAndStoreAccessToken(code){
  var nextURL = TOKEN_URL + '?client_id='+CLIENT_ID+'&client_secret='+CLIENT_SECRET+'&grant_type=authorization_code&redirect_uri='+REDIRECT_URL+'&code=' + code;
  
  var response = UrlFetchApp.fetch(nextURL).getContentText();   
  var tokenResponse = JSON.parse(response);
  
  //salesforce requires you to call against the instance URL (eg. https://na9.salesforce.com/)
  UserProperties.setProperty(baseURLPropertyName, tokenResponse.instance_url);
  UserProperties.setProperty(accessTokenPropertyName, tokenResponse.access_token);
  UserProperties.setProperty(refreshTokenPropertyName, tokenResponse.refresh_token);
  
  // return tokenResponse.refresh_token;
}

// Use the refresh token to set a new access token
function refreshToken() {
  var options = {
    "method" : "post"
  };
  var refreshToken = UserProperties.getProperty(refreshTokenPropertyName);
  var refreshURL = TOKEN_URL+'?grant_type=refresh_token&client_id='+CLIENT_ID+'&client_secret='+CLIENT_SECRET+'&refresh_token='+refreshToken;

  var response = UrlFetchApp.fetch(refreshURL, options);
  //Browser.msgBox("Refresh response code = " + response.getResponseCode());
  
  if (response.getResponseCode() == 200) {
    var tokenResponse = JSON.parse(response.getContentText());
    UserProperties.setProperty(baseURLPropertyName, tokenResponse.instance_url);
    UserProperties.setProperty(accessTokenPropertyName, tokenResponse.access_token);
    return true;
  } else {
    return false;
  }
}

// Send a GET request for data to Salesforce. 
// If it works, write the data to the spreadsheet.
// If it fails, try to refresh the token.
// If that fails, send the user back to the login screen.
function requestAndHandleData() {
  var response = getData();
  var responseCode = response.getResponseCode();
  if (responseCode < 300) {
    //Browser.msgBox(response.getContentText());
    writeDataToSpreadsheet(JSON.parse(response.getContentText()));
    
  } else if (responseCode == 401) { 
    // Try to reauthenticate using the refresh token. 
    // If that fails, try to log in.
    //Browser.msgBox("Error 401 - attempting to reauthenticate");
    if (refreshToken()) { requestAndHandleData(); } 
    else { login(); }
    
  } else { 
    // Some unexpected error - print it out
    Browser.msgBox("Error " + responseCode + ": " + response.getContentText());
  }
}

// Run the SOQL query. This is where we pass the access token to Salesforce.
// Returns an HTTPResponse object.
function getData() {
  var soql = 'SELECT+name,phone,industry+from+Account';
  var getDataURL = UserProperties.getProperty(baseURLPropertyName) + '/services/data/v26.0/query/?q='+soql;
  var response = UrlFetchApp.fetch(getDataURL,getUrlFetchOptions()); 
  return response;
}

// Create a JSON object containing options for Salesforce GET data request
function getUrlFetchOptions() {
  var token = UserProperties.getProperty(accessTokenPropertyName);
  return {
    "contentType" : "application/json",
    "muteHttpExceptions" : true, // Deal with response code directly rather than throwing an exception on invalid token
    "headers" : {
      "Authorization" : "Bearer " + token,
      "Accept" : "application/json"
    }
  };
}

// Process the JSON object returned by Salesforce after a successful query.
function writeDataToSpreadsheet(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  for (var i = 0; i < data.totalSize; i++) {
    var record = data.records[i];
    sheet.appendRow([record.Name, record.Phone, record.Industry]);
  }
}
