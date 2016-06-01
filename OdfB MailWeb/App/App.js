/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};

    // Common initialization function (to be called from each page)
    app.initialize = function () {    
    };

    app.redirectUri = "https://localhost:44308/app/auth/auth.html";
    app.tenant = "msp813992.onmicrosoft.com";
    app.clientId = "c87ccb91-ecb8-4586-abc0-20f180d39a6a";
    return app;
})();


angular
    .module('odfbMail', ['AdalAngular', 'officeuifabric.core', 'officeuifabric.components'])
    .config(function (adalAuthenticationServiceProvider, $httpProvider) {
        adalAuthenticationServiceProvider.init(
                {
                    // Config to specify endpoints and similar for your app
                    tenant: app.tenant,
                    clientId: app.clientId,
                    //localLoginUrl: "/login",  // optional
                    redirectUri: app.redirectUri,
                    //cacheLocation: 'localStorage',
                    endpoints: { "https://graph.microsoft.com": "https://graph.microsoft.com" }
                    //endpoints: endpoints  // If you need to send CORS api requests.
                },
                $httpProvider   // pass http provider to inject request interceptor to attach tokens
                );
    });


if (location.href.indexOf('access_token=') < 0) {
    Office.initialize = function () {
        console.log(">>> Office.initialize()");
        angular.bootstrap(document.getElementById('container'), ['odfbMail']);
    };
}
else {
    angular.bootstrap(document.getElementById('container'), ['odfbMail']);
}