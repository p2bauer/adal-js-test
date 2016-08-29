(function () {
  'use strict';

  var officeAddin = angular.module('officeAddin');

  officeAddin.config(['$httpProvider', 'adalAuthenticationServiceProvider', 'appId', 'addInHost', adalConfigurator]);

  function adalConfigurator($httpProvider, adalProvider, appId, addInHost) {
    
    var adalConfig = {
      tenant: 'common',
      clientId: appId,
      extraQueryParameter: 'nux=1',

      // this is a custom redirectUrl so we can easily see what's getting returned from AD without adal/angular in the way
      redirectUri: addInHost + '/appread/frameRedirect.html', 
      
      endpoints: {
        'https://graph.microsoft.com': 'https://graph.microsoft.com'
      }
      // cacheLocation: 'localStorage', // enable this for IE, as sessionStorage does not work for localhost. 
    };
    adalProvider.init(adalConfig, $httpProvider);
  }
})();