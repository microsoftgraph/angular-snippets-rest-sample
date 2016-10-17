 # Microsoft Graph Snippets Sample for AngularJS

The Microsoft Graph API exposes multiple APIs from Microsoft cloud services through a single REST API endpoint. This repository shows you how to access multiple resources, including Microsoft Azure Active Directory (AD) and the Office 365 APIs, by making HTTP requests to the Microsoft Graph API in an Angular application. 

![Microsoft Graph Snippets sample screenshot](./README assets/screenshot.jpg)

**Note: If possible, please use this sample with a "non-work" or test account in Office 365. The sample does not clean up the created objects in your mailbox, calendar, contacts, and objects created from additional operations. At this time you'll have to manually remove these artifacts, for example, sample mails, contacts, and calendar events.**  

## Prerequisites

* [Node.js](https://nodejs.org/). Node is required to run the sample on a development server and to install dependencies. 
* An Office 365 admin account. You can sign up for [an Office 365 Developer subscription](https://aka.ms/devprogramsignup) that includes the resources that you need to start building Office 365 apps.

     > Note: If you already have a subscription, the previous link sends you to a page with the message *Sorry, you can’t add that to your current account*. In that case use an account from your current Office 365 subscription.

## Register the app

1.	Sign in to the [Azure Management Portal](http://manage.windowsazure.cn) using your Azure AD credentials.
2.	Click **Active Directory** on the left menu, then select the directory for your Office 365 developer site.
3.	On the top menu, click **Applications**.
4.	Click **Add** from the bottom menu.
5.	On the **What do you want to do page**, click **Add an application my organization is developing**.
6.	On the **Tell us about your application page**, select **Web application and/or web API** for type and enter a friendly name for the application.
7.	Click the arrow icon on the lower-right corner of the page.
8.	On the **Application information** page, enter **http://127.0.0.1:8080/** for the sign-on and redirect URI values.
9.	Once the application is successfully added, you'll be taken to the **Quick Start** page for the application. From there, select **Configure** in the top menu.
10.	Under **permissions to other applications**, select **Add application**. In the dialog box, select the **Microsoft Graph** application. After you return to the application configuration page, select the following Delegated permissions:

   - Read and write access to user profile
   - Read all user' full profiles
   - Read and write directory data
   - Access directory data as the signed in user
   - Read user mail
   - Send mail as a user
   - Have full access to user calendars
   - Read user contacts
   - Have full access to user files and files shared with user

11.	Copy the value specified for **Client ID** on the **Configure** page.
12.	Enable your app to use the Implicit grant type.  
  a. In the bottom menu, choose **Manage Manifest** > **Download Manifest**.  
  b. Open the manifest file, and set the **oauth2AllowImplicitFlow** property to **true**.  
  c. In the management portal, choose **Manage Manifest** > **Upload Manifest**, and upload the updated manifest file.

13. Click **Save** in the bottom menu.

## Configure and run the app

1. Using your favorite IDE, open **config.js** in *public/scripts*.
2. Replace *ENTER_YOUR_APP_ID* with the application ID of your registered Azure application.
3. Install project dependencies with Node's package manager (npm) by running ```npm install``` in the project's root directory on the command line.
4. Start the development server by running ```node server.js``` in the project's root directory.
5. Navigate to ```http://localhost:8080/``` in your web browser and sign in to the app using your Office 365 admin credentials.

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

We'd love to get your feedback about the Microsoft Graph Snippets Sample for AngularJS. You can send your questions and suggestions to us in the [Issues](https://github.com/microsoftgraph/angular-snippets-rest-sample/issues) section of this repository.

Your feedback is important to us. Connect with us on [Stack Overflow](http://stackoverflow.com/questions/tagged/office365+or+microsoftgraph). Tag your questions with [MicrosoftGraph] and [office365].
  
## Additional resources

* [Office Dev Center](http://dev.office.com/)
* [Microsoft Graph API](http://graph.microsoft.io)
* [Microsoft Graph Snippets Sample for AngularJS](https://github.com/microsoftgraph/angular-connect-rest-sample)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
