/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

(function () {
	angular
		.module('app')
		.factory('usersFactory', usersFactory);

	usersFactory.$inject = ['$log', '$http', '$q', 'commonFactory'];
	function usersFactory($log, $http, $q, common) {
		var users = {}; 
 
		// Methods
		users.getMe = getMe;
		users.getUsers = getUsers;
		users.createUser = createUser;
		users.getDrive = getDrive;
		users.getEvents = getEvents;
		users.createEvent = createEvent;
		users.updateEvent = updateEvent;
		users.deleteEvent = deleteEvent;
		users.getMessages = getMessages;
		users.sendMessage = sendMessage;
		users.getUserPhoto = getUserPhoto;
		users.getManager = getManager;
		users.getDirectReports = getDirectReports;
		users.getMemberOf = getMemberOf;
		users.getFiles = getFiles;
		users.createFile = createFile;
		users.downloadFile = downloadFile;
		users.updateFile = updateFile;
		users.renameFile = renameFile;
		users.deleteFile = deleteFile;
		users.createFolder = createFolder;
		
		/////////////////////////////////////////
		// End of exposed properties and methods.
		
		var baseUrl = common.baseUrl;

		/**
		 * Get information about the signed-in user.
		 */
		function getMe(useSelect) {
			var reqUrl = baseUrl + '/me';

			// Append $select OData query string.
			if (useSelect) {
				reqUrl += '?$select=AboutMe,Responsibilities,Tags'
			}

			var req = {
				method: 'GET',
				url: reqUrl
			}

			return $q(function (resolve, reject) {
				$http(req)
					.then(function (res) {
						resolve(res);
					}, function (err) {
						resolve(err);
					});
			});
		};
		
		/**
		 * Get existing users collection from the tenant.
		 */
		function getUsers(useFilter) {
			var reqUrl = baseUrl + '/myOrganization/users';
			
			// Append $filter OData query string.
			if (useFilter) {
				// This filter will return users in your tenant based in the US.
				reqUrl += '?$filter=country eq \'United States\'';
			}

			var req = {
				method: 'GET',
				url: reqUrl
			};

			return $http(req);
		};
		
		/**
		 * Add a user to the tenant's users collection.
		 */
		function createUser(tenant) {
			var randomUserName = common.guid();
			
			// The data in newUser are the minimum required properties.
			var newUser = {
				accountEnabled: true,
				displayName: 'User ' + randomUserName,
				mailNickname: randomUserName,
				passwordProfile: {
					password: 'p@ssw0rd!'
				},
				userPrincipalName: randomUserName + '@' + tenant
			};

			var req = {
				method: 'POST',
				url: baseUrl + '/myOrganization/users',
				data: newUser
			};

			return $http(req);
		};
		
		/**
		 * Get the signed-in user's drive.
		 */
		function getDrive() {
			var req = {
				method: 'GET',
				url: baseUrl + '/me/drive'
			};

			return $http(req);
		};
		
		/**
		 * Get the signed-in user's calendar events.
		 */
		function getEvents() {
			var req = {
				method: 'GET',
				url: baseUrl + '/me/events'
			};

			return $http(req);
		};
		
		/**
		 * Adds an event to the signed-in user's calendar.
		 */
		function createEvent() {
			// The new event will be 30 minutes and take place tomorrow at the current time.
			var startTime = new Date();
			startTime.setDate(startTime.getDate() + 1);
			var endTime = new Date(startTime.getTime() + 30 * 60000);

			var newEvent = {
				Subject: 'Weekly Sync',
				Location: {
					DisplayName: 'Water cooler'
				},
				Attendees: [{
					Type: 'Required',
					EmailAddress: {
						Address: 'mara@fabrikam.com'
					}
				}],
				Start: {
					'DateTime': startTime,
					'TimeZone': 'PST'
				},
				End: {
					'DateTime': endTime,
					'TimeZone': 'PST'
				},
				Body: {
					Content: 'Status updates, blocking issues, and next steps.',
					ContentType: 'Text'
				}
			};

			var req = {
				method: 'POST',
				url: baseUrl + '/me/events',
				data: newEvent
			};

			return $http(req);
		};
		
		/**
		 * Creates an event, adds it to the signed-in user's calendar, and then
		 * updates the Subject.
		 */
		function updateEvent() {
			var deferred = $q.defer();

			var eventUpdates = {
				Subject: 'Sync of the Week'
			};
			
			// Create an event to update.
			createEvent()
			// If successful, take event ID and update it.
				.then(function (response) {
					var eventId = response.data.id;

					var req = {
						method: 'PATCH',
						url: baseUrl + '/me/events/' + eventId,
						data: eventUpdates
					};

					deferred.resolve($http(req));
				}, function (error) {
					deferred.reject({
						setupError: 'Unable to create an event to update.',
						response: error
					});
				});

			return deferred.promise;
		};
		
		/**
		 * Creates an event, adds it to the signed-in user's calendar, and then
		 * deletes the event.
		 */
		function deleteEvent() {
			var deferred = $q.defer();
			
			// Create an event to update first.
			createEvent()
			// If successful, take event ID and update it.
				.then(function (response) {
					var eventId = response.data.id;

					var req = {
						method: 'DELETE',
						url: baseUrl + '/me/events/' + eventId
					};

					deferred.resolve($http(req));
				}, function (error) {
					deferred.reject({
						setupError: 'Unable to create an event to delete.',
						response: error
					});
				});

			return deferred.promise;
		};

		/**
		 * Get the signed-in user's messages.
		 */
		function getMessages() {
			var req = {
				method: 'GET',
				url: baseUrl + '/me/messages'
			};

			return $http(req);
		};
		
		/**
		 * Send a message as the signed-in user.
		 */
		function sendMessage(recipientEmailAddress) {
			var newMessage = {
				Message: {
					Subject: 'Microsoft Graph snippets',
					Body: {
						ContentType: 'Text',
						Content: 'You can send an email by making a POST request to /me/microsoft.graph.sendMail.'
					},
					ToRecipients: [
						{
							EmailAddress: {
								Address: recipientEmailAddress
							}
						}
					]
				},
				SaveToSentItems: true
			};

			var req = {
				method: 'POST',
				url: baseUrl + '/me/microsoft.graph.sendMail',
				data: newMessage
			};

			return $http(req);
		};
		
		/**
		 * Get signed-in user's photo.
		 */
		function getUserPhoto() {
			var req = {
				method: 'GET',
				url: baseUrl + '/me/photo'
			};
			
			return $http(req);
		};
		
		/**
		 * Get signed-in user's manager.
		 */
		function getManager() {
			var req = {
				method: 'GET',
				url: baseUrl + '/me/manager'
			};
			
			return $http(req);
		};
		
		/**
		 * Get signed-in user's direct reports.
		 */
		function getDirectReports() {
			var req = {
				method: 'GET',
				url: baseUrl + '/me/directReports'
			};
			
			return $http(req);
		};
		
		/**
		 * Get groups that signed-in user is a member of.
		 */
		function getMemberOf() {
			var req = {
				method: 'GET',
				url: baseUrl + '/me/memberOf'
			};
			
			return $http(req);
		};
		
		/**
		 * Get signed-in user's files.
		 */
		function getFiles() {
			var req = {
				method: 'GET',
				url: baseUrl + '/me/drive/root/children'
			};
			
			return $http(req);
		};
		
		/**
		 * Create a file in signed-in user's root directory.
		 */
		function createFile() {
			var randomFileName = common.guid() + '.txt';
			
			var req = {
				method: 'PUT',
				url: baseUrl + '/me/drive/root/children/' + randomFileName + '/content',
				data: {
					content: randomFileName + ' is the name of this file.'
				}
			}
			
			return $http(req);
		};
		
		/**
		 * Get contents of a specific file.
		 */
		function downloadFile() {
			var deferred = $q.defer();
			
			createFile()
				.then(function (response) {
					var fileId = response.data.id;
					
					var req = {
						method: 'GET',
						url: baseUrl + '/me/drive/items/' + fileId + '/content'
					};

					deferred.resolve($http(req));
				}, function (error) {
					deferred.reject({
						setupError: 'Unable to create a file to download.',
						response: error
					});
				});
			
			return deferred.promise;
		};
		
		/**
		 * Updates the contents of a specific file.
		 */
		function updateFile() {
			var deferred = $q.defer();
			
			createFile()
				.then(function (response) {
					var fileId = response.data.id;
					
					var req = {
						method: 'PUT',
						url: baseUrl + '/me/drive/items/' + fileId + '/content',
						data: {
							content: 'Updated file contents.'
						}
					};

					deferred.resolve($http(req));
				}, function (error) {
					deferred.reject({
						setupError: 'Unable to create a file to update.',
						response: error
					});
				});
			
			return deferred.promise;
		};
		
		/**
		 * Renames a specific file.
		 */
		function renameFile(fileId) {
			var deferred = $q.defer();
			
			createFile()
				.then(function (response) {
					var fileId = response.data.id;
					var fileName = response.data.name.replace('.txt', '-renamed.txt');
					
					var req = {
						method: 'PATCH',
						url: baseUrl + '/me/drive/items/' + fileId,
						data: {
							name: fileName
						}
					};

					deferred.resolve($http(req));
				}, function (error) {
					deferred.reject({
						setupError: 'Unable to create a file to rename.',
						response: error
					});
				});
			
			return deferred.promise;
		};
		
		/**
		 * Deletes a specific file.
		 */
		function deleteFile(fileId) {
			var deferred = $q.defer();
			
			createFile()
				.then(function (response) {
					var fileId = response.data.id;
					
					var req = {
						method: 'DELETE',
						url: baseUrl + '/me/drive/items/' + fileId
					};

					deferred.resolve($http(req));
				}, function (error) {
					deferred.reject({
						setupError: 'Unable to create a file to delete.',
						response: error
					});
				});
			
			return deferred.promise;
		};
		
		/**
		 * Creates a folder in the root directory.
		 */
		function createFolder() {
			var req = {
				method: 'POST',
				url: baseUrl + '/me/drive/root/children',
				data: {
					name: common.guid(),
					folder: {},
					'@name.conflictBehavior': 'rename'
				}
			};
			
			return $http(req);
		};

		return users;
		}
})();
