/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

// This sample uses an open source OAuth 2.0 library that is compatible with the Azure AD v2.0 endpoint. 
// Microsoft does not provide fixes or direct support for this library. 
// Refer to the library’s repository to file issues or for other support. 
// For more information about auth libraries see: https://azure.microsoft.com/documentation/articles/active-directory-v2-libraries/ 
// Library repo: https://github.com/MrSwitch/hello.js

"use strict";

(function () {
  angular
    .module('app')
    .service('authHelper', ['$http', function ($http) {

      // Initialize the auth request.
      hello.init( {
        aad: appId // from public/scripts/config.js
        }, {
          response_type: 'token',
          redirect_uri: redirectUrl,
          scope: scopes + ' ' + adminScopes
        });
      
      // Set global headers for the request.
      function setAuthHeader(accessToken) {

          // Add the required Authorization header with bearer token.
          $http.defaults.headers.common.Authorization = 'Bearer ' + accessToken;
          
          // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
          $http.defaults.headers.common.SampleID = 'angular-snippets-rest-sample';
      }

      return {

        // Sign in and sign out the user.
        login: function login() {
          hello('aad').login({
            display: 'page',
            state: 'abcd'
          });
        },
        logout: function logout() {
          hello('aad').logout();
          delete localStorage.auth;
        },

        // Get a valid access token.
        getToken: function getToken() {

          // If the stored token is valid for another 5 minutes, we'll use it.
          let auth = angular.fromJson(localStorage.auth);
          let expiration = new Date();
          if (auth.access_token && expiration.setTime((auth.expires - 300) * 1000) > new Date()) { 
            setAuthHeader(auth.access_token);
          } else {

            // This sample just redirects the user to sign back in when the token expires.
            this.logout();
            this.login();
          } 
        }
      }
    }]);
})();