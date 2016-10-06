# Microsoft Graph Snippets Sample for AngularJS (REST)

## Table of contents

* [Prerequisites](#prerequisites)
* [Register the application](#register-the-application)
* [Build and run the sample](#build-and-run-the-sample)
* [Code of note](#code-of-note)
* [Questions and comments](#questions-and-comments)
* [Contributing](#contributing)
* [Additional resources](#additional-resources)

This sample shows how to use the Microsoft Graph API to send email, manage groups, and perform other activities. Microsoft Graph exposes multiple APIs from Microsoft cloud services through a single REST API endpoint. This repository shows you how to access multiple resources, including Microsoft Azure Active Directory (AD) and the Office 365 APIs, by making HTTP requests to Microsoft Graph in an AngularJS app. The sample uses the Azure AD v2.0 endpoint, which supports Microsoft Accounts and work or school Office 365 accounts.

![Microsoft Graph Snippets sample screenshot](./README assets/screenshot.jpg)

**Note:** This sample does not always clean up the entities that it creates, so you might want to use a test account to run the sample.

## Prerequisites

* [Node.js](https://nodejs.org/). Node is required to run the sample on a development server and to install dependencies. 

* [Bower](https://bower.io). Bower is required to install front-end dependencies.

* Either a [Microsoft account](https://www.outlook.com) or [work or school account](http://dev.office.com/devprogram) (admin)

## Register the application

1. Sign into the [App Registration Portal](https://apps.dev.microsoft.com/) using either your personal or work or school account.

2. Choose **Add an app**.

3. Enter a name for the app, and choose **Create application**.
	
	The registration page displays, listing the properties of your app.
 
4. Copy the application ID. This is the unique identifier for your app that you'll use to configure the sample.

5. Under **Platforms**, choose **Add Platform** > **Web**.

6. Make sure the Allow Implicit Flow check box is selected, and enter http://localhost:8080 as the Redirect URI. 

7. Choose **Save**.

## Build and run the sample

1. Download or clone the Microsoft Graph Snippets Sample for AngularJS.

2. Using your favorite IDE, open **config.js** in *public/scripts*.

3. Replace the **appId** placeholder value with the application ID of your registered Azure application.

4. In a command prompt, run the following commands in the sample's root directory. This installs project dependencies, including the [HelloJS](http://adodson.com/hello.js/) client-side authentication library.

  ```
npm install
bower install hello
  ```
  
5. Run `npm start` to start the development server.

6. Navigate to `http://localhost:8080` in your web browser.

7. Sign in with your personal or admin work or school account and grant the requested permissions.

8. Choose a snippet from the left-hand navigation pane, and then choose the **Run snippet** button. The request and response display in the center pane.


### How the sample affects your data

This sample runs REST commands that create, read, update, or delete data. The sample creates fake entities so that your actual tenant data is unaffected. The sample will leave behind the fake entities that it creates.

<a name="contributing"></a>
## Contributing ##

If you'd like to contribute to this sample, see [CONTRIBUTING.MD](/CONTRIBUTING.md).

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Questions and comments

We'd love to get your feedback about the Microsoft Graph API snippets with Angular sample. You can send your questions and suggestions to us in the [Issues](https://github.com/microsoftgraph/angular-snippets-rest-sample/issues) section of this repository.

Your feedback is important to us. Connect with us on [Stack Overflow](http://stackoverflow.com/questions/tagged/microsoftgraph). Tag your questions with [MicrosoftGraph].
  
## Additional resources

* [Microsoft Graph](http://graph.microsoft.io)
* [Other Microsoft Graph samples for AngularJS](https://github.com/microsoftgraph?utf8=%E2%9C%93&query=angular)

## Copyright
Copyright (c) 2016 Microsoft. All rights reserved.
