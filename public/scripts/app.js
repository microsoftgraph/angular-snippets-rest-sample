/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

(function () {
  angular
    .module('app', [
      'ngRoute',
      'ui.bootstrap',
      'AdalAngular',
      'angular-loading-bar',
      'ladda',
      'LocalStorageModule'
    ])
    .config(config);

  function config($routeProvider, adalAuthenticationServiceProvider, $httpProvider, cfpLoadingBarProvider, localStorageServiceProvider) {
    // Configure the routes. 
    $routeProvider
      .when('/', {
        templateUrl: 'views/main.html',
        controller: 'MainController',
        controllerAs: 'main'
      })
      .otherwise({
        redirectTo: '/'
      });
      
    // Allow cross domain requests to be made.
    $httpProvider.defaults.useXDomain = true;
    delete $httpProvider.defaults.headers.common['X-Requested-With'];
      
    // Configure ADAL JS. 
    adalAuthenticationServiceProvider.init(
      {
        instance: 'https://login.chinacloudapi.cn/',
        clientId: clientId,
        endpoints: {
          'https://microsoftgraph.chinacloudapi.cn': 'https://microsoftgraph.chinacloudapi.cn'
        }
      },
      $httpProvider
    );
    
    // Remove spinner from loading bar.
    cfpLoadingBarProvider.includeSpinner = false;
    
    // Local storage configuration.
    localStorageServiceProvider
      .setPrefix('unifiedApiSnippets');
  };
})();
