# [ARCHIVED] Office 365 People Controls

========================

**Note:** This repo is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the Office/SharePoint development workload in Visual Studio, or the [Yeoman generator for[Arch Office Add-ins](https://github.com/OfficeDev/generator-office), and follow security best practices as you develop the add-in." The Office 365 People Controls provide a simple but extensible way to add standard Office experience and access data from Office 365 service. All the controls are developed in JavaScript to provide universal compatibility without the additional overhead of other frameworks. We build the controls with two parts, Web UI and Data Provider, so that developers could customize the control easily based on the interface we defined.

Currently, we provide two basic people controls
 - People Picker
 - Persona Card

You could find more description from [Office 365 JavaScript controls](https://msdn.microsoft.com/en-us/office/office365/howto/javascript-controls)

## Web UI
The standard Office 365 web experience comes from Office UI Fabric, you could also visit their GitHub repository - [OfficeDev/Office-UI-Fabric](https://github.com/OfficeDev/Office-UI-Fabric)


## Data Provider
We provide a sample data provider which retrieves data from Office 365, you could get more detail about how to access Office 365 data from - [Office 365 API reference](https://msdn.microsoft.com/office/office365/HowTo/rest-api-overview).

If you want to use the sample provider, please remember to type in your Office 365 client ID before initialize the provider. 

Here are the key steps
 - Create your own Office 365 client ID
 - Set permissions for the API you plan to use
 - Set "Redirect URI" for your web app
 - Configure "Implicit Grant" for OAuth flow

For more detail guidance, you could check from - [Create an app with Office 365 APIs](https://msdn.microsoft.com/office/office365/howto/undefined/office/office365/howto/getting-started-Office-365-APIs)

## Permissions 
You need to configure permissions for your Office 365 app based on the API and scope you want to access. 

Here are the permissions sample data provider requires

|Feature|Application Name|Delegated Permission|Comments|
|:-----|:-----|:-----|:-----|
|Login|Azure Active Directory|Sign in and read user profile||
|People Picker|Azure Active Directory|Read all users' basic profiles||	
|Persona Card|Azure Active Directory|Read all users' basic profiles||	
|Persona Card|Azure Active Directory|Read all users' full profiles|AAD requires higher permission if app needs to get users' phone numbers. "admin_consent" need to be set for parameter "prompt" in [authorization code grant flow](https://msdn.microsoft.com/en-us/library/azure/dn645542.aspx).|

## License
 - All files on the Office 365 People Controls repository are subject to the MIT license. Please read the License 

file at the root of the project. 
 - All the Web UI are based on [OfficeDev/Office-UI-Fabric](https://github.com/OfficeDev/Office-UI-Fabric)
 - Usage of the fonts referenced on Office UI Fabric files is subject to the terms listed here 

## Sample Site
We provide a sample site in the "example" folder. In this site, you could find
 - Demo for controls
 - API guidance
 - Test page

Here are the key steps for running your own sample site

install nodejs
 - http://nodejs.org

To install development packages - From the root of your local git repository
 - npm install gulp -g
 - npm install
 - npm install dev

To build minified files
 - gulp
Then you could deploy to your web service

Or you could run a local copy of the sample by
 - gulp
 - node server.js
 - navigate to http://localhost:3000/
