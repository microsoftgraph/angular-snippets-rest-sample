/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

(function() {
	angular
		.module('app')
		.factory('contactsFactory', contactsFactory);
		
	contactsFactory.$inject = ['$log', '$http', '$q', 'commonFactory'];
	function contactsFactory($log, $http, $q, common) {
		var contacts = {};
		
		// Methods
		contacts.getContacts = getContacts;
		
		/////////////////////////////////////////
		// End of exposed properties and methods.
				
		var baseUrl = common.baseUrl;
				
		/**
		 * Gets the list of contacts that exist in the tenant.
		 */
		function getContacts() {
			var req = {
				method: 'GET',
				url: baseUrl + '/myOrganization/contacts'
			};

			return $http(req);
		};
				
		return contacts;
	};
})();

// *********************************************************
//
// O365-Angular-Unified-API-Snippets, https://github.com/OfficeDev/O365-Angular-Unified-API-Snippets
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