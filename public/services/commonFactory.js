/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
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