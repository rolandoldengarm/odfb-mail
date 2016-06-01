/// <reference path="../App.js" />

angular
    .module('odfbMail')
    .controller('homeCtrl', function ($scope, $http) {
        $scope.files = [];
        $scope.retrievingData = false;
        $scope.loginStatus = "Disconnected";
        //Office.initialize = function (reason) {
        //    $(document).ready(function () {
        //        app.initialize();                
        //    });
        //};
        
        $scope.signIn = function () {
            
            var url = "https://localhost:44308/app/auth/auth.html";

            Office.context.ui.displayDialogAsync(url, { height: 40, width: 40, requireHTTPS: true }, function (result) {
                _dlg = result.value;
                _dlg.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, function (authResult) {                                                          
                    
                    if (authResult.message === "success") {
                        $scope.loginStatus = "Connected";
                        $scope.getFiles();
                    }
                    else {
                        $scope.loginStatus = "Authentication Failed";
                    }
                        
                    _dlg.close();
                });
            });
        }

        $scope.getFiles = function () {
            if ($scope.retrievingData) return;
            $scope.retrievingData = true;
            var from = Office.context.mailbox.item.from;

            $http.get('https://graph.microsoft.com/v1.0/users/' + from.emailAddress + '/drive/root/children').then(function (files) {
                $scope.retrievingData = false;
                $scope.files = files.data.value;
            }, function (err) {
                $scope.retrievingData = false;
                console.log(err);
            }
            );
        }
});