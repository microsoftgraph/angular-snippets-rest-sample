/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

(function() {
	angular
		.module('app')
		.factory('drivesFactory', drivesFactory);
		
	drivesFactory.$inject = ['$log', '$http', '$q', 'commonFactory'];
	function drivesFactory($log, $http, $q, common) {
		var drives = {};
		
		// Methods
		drives.getDrives = getDrives;
		
		/////////////////////////////////////////
		// End of exposed properties and methods.
				
		var baseUrl = common.baseUrl;
				
		/**
		 * Gets the list of drives that exist in the tenant.
		 */
		function getDrives() {
			var req = {
				method: 'GET',
				url: baseUrl + '/myOrganization/drives'
			};

			return $http(req);
		};
				
		return drives;
	};
})(); 
