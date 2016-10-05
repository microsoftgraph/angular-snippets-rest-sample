# Microsoft Graph snippets with Angular

The Microsoft Graph API exposes multiple APIs from Microsoft cloud services through a single REST API endpoint. This repository shows you how to access multiple resources, including Microsoft Azure Active Directory (AD) and the Office 365 APIs, by making HTTP requests to the Microsoft Graph API in an Angular application. 

![Office 365 Angular Snippets sample screenshot](./README assets/screenshot.jpg)

**Note: If possible, please use this sample with a "non-work" or test account in Office 365. The sample does not clean up the created objects in your mailbox, calendar, contacts, and objects created from additional operations. At this time you'll have to manually remove these artifacts, for example, sample mails, contacts, and calendar events.**  



## Prerequisites

* [Node.js](https://nodejs.org/). Node is required to run the sample on a development server and to install dependencies. 
* An Office 365 admin account. You can sign up for [an Office 365 Developer subscription](https://aka.ms/devprogramsignup) that includes the resources that you need to start building Office 365 apps.

     > Note: If you already have a subscription, the previous link sends you to a page with the message *Sorry, you can’t add that to your current account*. In that case use an account from your current Office 365 subscription.

## Register the app

1. Sign in to the [Azure portal](https://portal.azure.com/).
2. On the top bar, click on your account and under the **Directory** list, choose the Active Directory tenant where you wish to register your application.
3. Click on **More Services** in the left hand nav, and choose **Azure Active Directory**.
4. Click on **App registrations** and choose **Add**.
5. Enter a friendly name for the application, for example 'MSGraphSnippetsAngular' and select 'Web app/API' as the **Application Type**. For the Sign-on URL, enter *http://127.0.0.1:8080/*. Click on **Create** to create the application.
6. While still in the Azure portal, choose your application, click on **Settings** and choose **Properties**.
7. Find the Application ID value and copy it to the clipboard.
8. Enable your app to use the Implicit grant type.
  a. Choose the **Manifest** tab above the app details.
  b. Choose **Edit**, and then set the **oauth2AllowImplicitFlow** property to **true**.
  c. Choose **Save**.
9. Configure Permissions for your application:
  a. In the **Settings** menu, choose the **Required permissions** section, click on **Add**, then **Select an API**, and select **Microsoft Graph**.
  b. Click on Select Permissions. Select the following delegated permissions and then choose **Done**.

   - Read and write access to user profile
   - Read all user' full profiles
   - Read and write directory data
   - Access directory data as the signed in user
   - Read user mail
   - Send mail as user
   - Have full access to user calendars
   - Read user contacts
   - Have full access to user files

## Configure and run the app

1. Using your favorite IDE, open **config.js** in *public/scripts*.
2. Replace *{your_app_client_ID}* with the client ID of your registered Azure application.
3. Install project dependencies with Node's package manager (npm) by running ```npm install``` in the project's root directory on the command line.
4. Start the development server by running ```node server.js``` in the project's root directory.
5. Navigate to ```http://127.0.0.1:8080/``` in your web browser and sign in to the app using your Office 365 admin credentials.

## Understanding the app

This sample demonstrates several concepts including:

* Learn how to make REST calls that target data stored in Office 365 (including Exchange Online and SharePoint Online).
* Understand how to use the [Azure Active Directory Authentication Library (ADAL) for JavaScript](https://github.com/AzureAD/azure-activedirectory-library-for-js).

### Connecting to Office 365

The sample provides the code required to display the Office 365 sign in page if there are no tokens available already in the local cache. The sample uses the Azure Active Directory Library for JavaScript to manage the tokens required for the app to use Office 365 services.

The code that uses ADAL JS to authenticate with Azure AD is located in the following files (located in *public/*):

* *scripts/app.js* - The service is configured here.
* *scripts/controllers/navbarController.js* - The login/logOut methods exposed by ADAL are leveraged here.

### Snippets

This sample is a repository of Microsoft Graph API code snippets that demonstrate how to work with Office 365 objects like mail, calendar, contacts, files, and user profile information.

Each tenant-level resource collection has its own factory defined in *public/services* where you can see how the requests are constructed. Then the factories are called by *mainController.js* found in the *public/controllers* directory.

<a name="contributing"></a>
## Contributing ##

If you'd like to contribute to this sample, see [CONTRIBUTING.MD](/CONTRIBUTING.md).

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Questions and comments

We'd love to get your feedback about the Microsoft Graph API snippets with Angular sample. You can send your questions and suggestions to us in the [Issues](https://github.com/microsoftgraph/angular-snippets-rest-sample/issues) section of this repository.

Your feedback is important to us. Connect with us on [Stack Overflow](http://stackoverflow.com/questions/tagged/office365+or+microsoftgraph). Tag your questions with [MicrosoftGraph] and [office365].
  
## Additional resources

* [Office Dev Center](http://dev.office.com/)
* [Microsoft Graph API](http://graph.microsoft.io)
* [Office 365 Angular Connect sample using Microsoft Graph API](https://github.com/microsoftgraph/angular-connect-rest-sample)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
