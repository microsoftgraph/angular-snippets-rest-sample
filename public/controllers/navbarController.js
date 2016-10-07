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
	NavbarController.$inject = ['$log', '$scope', 'adalAuthenticationService'];
	function NavbarController($log, $scope, adalAuthenticationService) {
		var vm = this;
		
		// Properties
		vm.isCollapsed;
		vm.isopen = false;	
		
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
		})(); 
		
		/**
		 * Expose the login method to the view.
		 */
		function connect() {
			$log.debug('Connecting to Office 365...');
			adalAuthenticationService.login();
		};
		
		/**
		 * Expose the logOut method to the view.
		 */
		function disconnect() {
			$log.debug('Disconnecting from Office 365...');
			adalAuthenticationService.logOut();
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
