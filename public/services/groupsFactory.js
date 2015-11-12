/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

(function () {
	angular
		.module('app')
		.factory('groupsFactory', groupsFactory);

	groupsFactory.$inject = ['$log', '$http', '$q', 'commonFactory'];
	function groupsFactory($log, $http, $q, common) {
		var groups = {}; 
		
		// Methods
		groups.getGroups = getGroups;
		groups.createGroup = createGroup;
		groups.getGroup = getGroup;
		groups.updateGroup = updateGroup;
		groups.deleteGroup = deleteGroup;
		groups.getMembers = getMembers;
		groups.getOwners = getOwners;
		
		/////////////////////////////////////////
		// End of exposed properties and methods.
		
		var baseUrl = common.baseUrl;
		
		/**
		 * Gets the list of groups that exist in the tenant.
		 */
		function getGroups() {
			var req = {
				method: 'GET',
				url: baseUrl + '/myOrganization/groups'
			};

			return $http(req);
		};
		
		/**
		 * Creates a new security group in the tenant.
		 */
		function createGroup() {
			var uuid = common.guid();

			var newGroup = {
				displayName: uuid,
				mailEnabled: false, // Must be false. This is because only pure security groups can be created using the Microsoft Graph API.
				mailNickname: uuid,
				securityEnabled: true // Must be true. This is because only pure security groups can be created using the Microsoft Graph API.
			};

			var req = {
				method: 'POST',
				url: baseUrl + '/myOrganization/groups',
				data: newGroup
			};

			return $http(req);
		};
		
		/**
		 * Gets information about a specific group.
		 */
		function getGroup() {
			return getGroupsSetup('');
		};
		
		/**
		 * Updates the description of a specific group.
		 */
		function updateGroup() {
			var deferred = $q.defer();

			// You can only update groups created via the Microsoft Graph API, so to make sure we have one,
			// we'll create it here and then update its description. 
			createGroup()
				.then(function (response) {
					var groupId = response.data.objectId;
					$log.debug('Group "' + groupId + '" was created. Updating the group\'s description...');

					var groupUpdates = {
						description: 'This is a group.'
					};

					var req = {
						method: 'PATCH',
						url: baseUrl + '/myOrganization/groups/' + groupId,
						data: groupUpdates
					};

					deferred.resolve($http(req));
				}, function (error) {
					deferred.reject({
						setupError: 'Unable to create a new group to update.',
						response: error
					});
				});

			return deferred.promise;
		};
		
		/**
		 * Deletes a specific group.
		 */
		function deleteGroup() {
			var deferred = $q.defer();

			// You can only delete groups created via the Microsoft Graph API, so to make sure we have one,
			// we'll create it here and then delete its description. 
			createGroup()
				.then(function (response) {
					var groupId = response.data.objectId;
					$log.debug('Group "' + groupId + '" was created. Deleting the group...');

					var req = {
						method: 'DELETE',
						url: baseUrl + '/myOrganization/groups/' + groupId
					};

					deferred.resolve($http(req));
				}, function (error) {
					deferred.reject({
						setupError: 'Unable to create a new group to delete.',
						response: error
					});
				});

			return deferred.promise;
		};
		
		/**
		 * Gets members of a specific group.
		 */
		function getMembers() {
			return getGroupsSetup('members');
		};

		/**
		 * Gets owners of a specific group.
		 */
		function getOwners() {
			return getGroupsSetup('owners');
		};

		/**
		 * Several snippets require a group ID to work. This method does the setup work
		 * required to get the ID of a group in the tenant and then makes a request to the
		 * desired navigation property. 
		 */
		function getGroupsSetup(endpoint) {
			var deferred = $q.defer();

			getGroups()
				.then(function (response) {
					// Check to make sure at least 1 group is returned.
					if (response.data.value.length >= 1) {
						var groupId = response.data.value[0].objectId;
						
						var req = {
							method: 'GET',
							url: baseUrl + '/myOrganization/groups/' + groupId + '/' + endpoint
						};

						deferred.resolve($http(req));
					}
					else {
						deferred.reject({
							setupError: 'Tenant doesn\'t have any groups.',
							response: response
						});
					}
				}, function (error) {
					deferred.reject({
						setupError: 'Unable to get list of tenant\'s groups.',
						response: error
					});
				});

			return deferred.promise;
		};

		return groups;
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