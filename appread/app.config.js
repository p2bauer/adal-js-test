(function () {
	var officeAddin = angular.module('officeAddin');

	// this is the AzureAD App ID for https://test.colligoapp.com :-)
	officeAddin.constant('appId', '3b154108-ca2f-4c5c-8aca-4466e141dfd9');

	// we save the current url so we can assemble redirect urls later!
	officeAddin.constant('addInHost', window.location.origin);
})();