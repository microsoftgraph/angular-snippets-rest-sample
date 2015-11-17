/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

(function () {
	angular
		.module('app')
		.factory('commonFactory', commonFactory);

	function commonFactory($http, $q) {
		var common = {};
		
		// Object constructors
		common.Snippet = Snippet;
		
		// Properties
		common.baseUrl = 'https://graph.microsoft.com/v1.0';
		
		// Methods
		common.guid = guid;
		
		/////////////////////////////////////////
		// End of exposed properties and methods.
		
		/**
		 * Constructor for the snippet object.
		 */
		function Snippet(title, description, documentationUrl, apiUrl, requireAdmin, run) {
			this.title = title;
			this.description = description;
			this.documentationUrl = documentationUrl;
			this.apiUrl = apiUrl;
			this.requireAdmin = requireAdmin;
			this.run = run;
		};
		
		/**
		 * Random GUID generator. Copied from Stack Overflow user "broofa".
		 * http://stackoverflow.com/users/109538/broofa
		 * http://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript
		 */
		function guid() {
			function s4() {
				return Math.floor((1 + Math.random()) * 0x10000)
					.toString(16)
					.substring(1);
			}
			return s4() + s4() + '-' + s4() + '-' + s4() + '-' +
				s4() + '-' + s4() + s4() + s4();
		};

		return common;
	}
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