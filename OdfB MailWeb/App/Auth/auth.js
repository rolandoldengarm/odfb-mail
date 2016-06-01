angular
    .module('odfbMail')
    .controller('authCtrl', function ($scope, adalAuthenticationService) {
        
        $scope.onLoginSuccess = $scope.$on("adal:loginSuccess", function () {
            
            Office.initialize = function (reason) {
                if (typeof Office.context.ui !== "undefined")
                    Office.context.ui.messageParent("success");
            };
            $scope.onLoginSuccess();
        }, true);

        $scope.onLoginFailure = $scope.$on("adal:loginFailure", function () {
            Office.initialize = function (reason) {
                if (typeof Office.context.ui !== "undefined")
                    Office.context.ui.messageParent("error");
            };
            $scope.onLoginFailure();
        });

        // optional
        $scope.onNotAuthorized = $scope.$on("adal:notAuthorized", function (event, rejection, forResource) {
            Office.initialize = function (reason) {
                if (typeof Office.context.ui !== "undefined")
                    Office.context.ui.messageParent("error");
            };
            $scope.onNotAuthorized();
        });
        $scope.signIn = function () {
           
            var user = adalAuthenticationService.getCachedToken();

            // Check if the user is cached
            if (!user) {
                adalAuthenticationService.login();
            }

            var response = { "status": "none", "accessToken": "" };            
        }
        $scope.signIn();
        //Office.initialize = function (reason) {
        //    $(document).ready($scope.signIn);
        //};
    });