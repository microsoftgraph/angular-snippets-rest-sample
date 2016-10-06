/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

(function () {
	angular
		.module('app')
		.controller('NavbarController', NavbarController);
		
	/**
	 * The NavbarController code.
	 */
	NavbarController.$inject = ['$log', '$scope', 'authHelper'];
	function NavbarController($log, $scope, authHelper) {
		var vm = this;
		
		// Properties
		vm.isCollapsed;
		vm.isopen = false;
		vm.isConnected;
		
		// Methods
		vm.connect = connect;
		vm.disconnect = disconnect;
		
		/////////////////////////////////////////
		// End of exposed properties and methods.
		
		/**
		 * This function does any initialization work the 
		 * controller needs.
		 */
		(function activate() {
			vm.isCollapsed = true;
			if (hello('aad').getAuthResponse()) vm.isConnected = true;
		})(); 
		
		/**
		 * Expose the login method to the view.
		 */
		function connect() {
			$log.debug('Connecting to Microsoft Graph.');
			authHelper.login();
		};

		/**
		 * Expose the logout method to the view.
		 */
		function disconnect() {
			$log.debug('Disconnecting from Microsoft Graph.');
			vm.isConnected = false;
			authHelper.logout();
		};

		/**
		 * Event listener for dropdown menu in navbar. 
		 */
		$scope.toggleDropdown = function ($event) {
			$event.preventDefault();
			$event.stopPropagation();
			$scope.status.isopen = !$scope.status.isopen;
		};
	};
})();
