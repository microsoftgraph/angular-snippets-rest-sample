/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
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

// *********************************************************
//
// O365-Angular-Microsoft-Graph-Snippets, https://github.com/OfficeDev/O365-Angular-Microsoft-Graph-Snippets
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************
