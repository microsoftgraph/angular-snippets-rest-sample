/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

(function () {
	angular
		.module('app')
		.controller('MainController', MainController);

	/**
	 * The MainController code.
	 */
	MainController.$inject = ['$scope', '$q', 'adalAuthenticationService', 'commonFactory', 'usersFactory', 'groupsFactory', 'drivesFactory'];
	function MainController($scope, $q, adalAuthenticationService, common, users, groups, drives) {
		var vm = this;
		
		// Snippet constructor from commonFactory.
		var Snippet = common.Snippet;
		
		/////////////////////////
		// Snippet             //
		// -------             //
		// Title               //
		// Description         //
		// Documenation URL    //
		// API URL             //
		// Require admin?      //
		// Snippet code        //
		/////////////////////////
		
		////////////////////////////////////////////////
		// All of the snippets that fall under the    //
		// 'users' tenant-level resource collection.  //
		////////////////////////////////////////////////
		var usersSnippets = {
			groupTitle: 'users',
			snippets: [
				///////////////////////////////
				//       USER SNIPPETS       // 
				///////////////////////////////				
				new Snippet(
					'GET myOrganization/users',
					'Gets all of the users in your tenant\'s directory.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list',
					common.baseUrl + '/myOrganization/users',
					false,				
					function () {
						doSnippet(partial(users.getUsers, false));
					}),
				new Snippet(
					'GET myOrganization/users?$filter=country eq \'United States\'',
					'Gets all of the users in your tenant\'s directory who are from the United States, using $filter.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list',
					common.baseUrl + '/myOrganization/users?$filter=country eq \'United States\'',
					false,
					function () {
						doSnippet(partial(users.getUsers, true));
					}),
				new Snippet(
					'POST myOrganization/users',
					'Adds a new user to the tenant\'s directory.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/user_post_users',
					common.baseUrl + '/myOrganization/users',
					true,					
					function () {
						doSnippet(partial(users.createUser, tenant));
					}),
				new Snippet(
					'GET me',
					'Gets information about the signed-in user.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/user_get',
					common.baseUrl + '/me',
					false,					
					function () {
						doSnippet(partial(users.getMe, false));
					}),
				new Snippet(
					'GET me?$select=AboutMe,Responsibilities,Tags',
					'Gets select information about the signed-in user, using $select.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/user_get',
					common.baseUrl + '/me?$select=AboutMe,Responsibilities,Tags',	
					false,				
					function () {
						doSnippet(partial(users.getMe, true));
					}),
				//////////////////////////////////////
				//        USER/DRIVE SNIPPETS       //
				//////////////////////////////////////								
				new Snippet(
					'GET me/drive',
					'Gets the signed-in user\'s drive.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/drive_get',
					common.baseUrl + '/me/drive',
					false,					
					function () {
						doSnippet(users.getDrive);
					}),
				///////////////////////////////////////
				//        USER/EVENTS SNIPPETS       //
				///////////////////////////////////////
				new Snippet(
					'GET me/events',
					'Gets the signed-in user\'s calendar events.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list_events',
					common.baseUrl + '/me/events',	
					false,				
					function () {
						doSnippet(users.getEvents);
					}),
				new Snippet(
					'POST me/events',
					'Adds an event to the signed-in user\'s calendar.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/user_post_events',
					common.baseUrl + '/me/events',	
					false,				
					function () {
						doSnippet(users.createEvent);
					}),
				new Snippet(
					'PATCH me/events/{Event.id}',
					'Adds an event to the signed-in user\'s calendar, then updates the subject of the event.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/event_update',
					common.baseUrl + '/me/events/{Event.id}',
					false,					
					function () {
						doSnippet(users.updateEvent);
					}),
				new Snippet(
					'DELETE me/events/{Event.id}',
					'Adds an event to the signed-in user\'s calendar, then deletes the event.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/event_delete',
					common.baseUrl + '/me/events/{Event.id}',
					false,					
					function () {
						doSnippet(users.deleteEvent);
					}),
				//////////////////////////////////////////
				//        USER/MESSAGES SNIPPETS        //
				//////////////////////////////////////////
				new Snippet(
					'GET me/messages',
					'Gets the signed-in user\'s emails.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list_messages',
					common.baseUrl + '/me/messages',	
					false,				
					function () {
						doSnippet(users.getMessages);
					}),
				new Snippet(
					'POST me/microsoft.graph.sendMail',
					'Sends an email as the signed-in user and saves a copy to their Sent Items folder.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/user_sendmail',
					common.baseUrl + '/me/microsoft.graph.sendMail',
					false,					
					function () {
						doSnippet(partial(users.sendMessage, adalAuthenticationService.userInfo.userName));
					}),
				//////////////////////////////////////////
				//          USER/FILES SNIPPETS         //
				//////////////////////////////////////////
				new Snippet(
					'GET me/drive/root/children',
					'Gets files from the signed-in user\'s root directory.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/item_list_children',
					common.baseUrl + '/me/drive/root/children',	
					false,				
					function () {
						doSnippet(users.getFiles);
					}),
				new Snippet(
					'PUT me/drive/root/children/{FileName}/content',
					'Creates a file with content in the signed-in user\'s root directory.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/item_uploadcontent',
					common.baseUrl + '/me/drive/root/children/{FileName}/content',	
					false,				
					function () {
						doSnippet(users.createFile);
					}),
				new Snippet(
					'GET me/drive/items/{File.id}/content',
					'Downloads a file.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/item_downloadcontent',
					common.baseUrl + '/me/drive/items/{File.id}/content',	
					false,				
					function () {
						doSnippet(users.downloadFile);
					}),
				new Snippet(
					'PUT me/drive/items/{File.id}/content',
					'Updates the contents of a file.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/item_downloadcontent',
					common.baseUrl + '/me/drive/items/{File.id}/content',	
					false,				
					function () {
						doSnippet(users.updateFile);
					}),
				new Snippet(
					'PATCH me/drive/items/{File.id}',
					'Renames a file.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/item_update',
					common.baseUrl + '/me/drive/items/{File.id}',	
					false,				
					function () {
						doSnippet(users.renameFile);
					}),
				new Snippet(
					'DELETE me/drive/items/{File.id}',
					'Deletes a file.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/item_delete',
					common.baseUrl + '/me/drive/items/{File.id}',	
					false,				
					function () {
						doSnippet(users.deleteFile);
					}),
				new Snippet(
					'POST me/drive/root/children',
					'Creates a folder in the signed-in user\'s root directory.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/item_post_children',
					common.baseUrl + '/me/drive/root/children',	
					false,				
					function () {
						doSnippet(users.createFolder);
					}),
				///////////////////////////////////////////////
				//        MISCELLANEOUS USER SNIPPETS        //
				///////////////////////////////////////////////
				new Snippet(
					'GET me/manager',
					'Gets the signed-in user\'s manager.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list_manager',
					common.baseUrl + '/me/manager',
					false,					
					function () {
						doSnippet(users.getManager);
					}),
				new Snippet(
					'GET me/directReports',
					'Gets the signed-in user\'s direct reports.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list_directreports',
					common.baseUrl + '/me/directReports',
					false,					
					function () {
						doSnippet(users.getDirectReports);
					}),
				new Snippet(
					'GET me/photo',
					'Gets the signed-in user\'s photo.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/profilephoto_get',
					common.baseUrl + '/me/photo',
					false,					
					function () {
						doSnippet(users.getUserPhoto);
					}),
				new Snippet(
					'GET me/memberOf',
					'Gets the groups that the signed-in user is a member of.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list_memberof',
					common.baseUrl + '/me/memberOf',	
					false,			
					function () {
						doSnippet(users.getMemberOf);
					})
			]
		};
		
		////////////////////////////////////////////////
		// All of the snippets that fall under the    //
		// 'groups' tenant-level resource collection. //
		////////////////////////////////////////////////
		var groupsSnippets = {
			groupTitle: 'groups',
			snippets: [
				/////////////////////////////////
				//       GROUPS SNIPPETS       // 
				/////////////////////////////////				
				new Snippet(
					'GET myOrganization/groups',
					'Gets all of the groups in your tenant\'s directory.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/group_list',
					common.baseUrl + '/myOrganization/groups',	
					false,				
					function () {
						doSnippet(groups.getGroups);
					}),
				new Snippet(
					'POST myOrganization/groups',
					'Adds a new security group to the tenant.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/group_post_groups',
					common.baseUrl + '/myOrganization/groups',	
					false,				
					function () {
						doSnippet(groups.createGroup);
					}),
				new Snippet(
					'GET myOrganization/groups/{Group.id}',
					'Gets information about a group in the tenant by ID.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/group_get',
					common.baseUrl + '/myOrganization/groups/{Group.id}',
					false,					
					function () {
						doSnippet(groups.getGroup);
					}),
				new Snippet(
					'PATCH myOrganization/groups/{Group.id}',
					'Adds a new group to the tenant, then updates the description of that group.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/group_update',
					common.baseUrl + '/myOrganization/groups/{Group.id}',
					false,				
					function () {
						doSnippet(groups.updateGroup);
					}),
				new Snippet(
					'DELETE myOrganization/groups/{Group.id}',
					'Adds a new group to the tenant, then deletes the group.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/group_delete',
					common.baseUrl + '/myOrganization/groups/{Group.id}',
					false,				
					function () {
						doSnippet(groups.deleteGroup);
					}),
				new Snippet(
					'GET myOrganization/groups/{Group.id}/members',
					'Gets the members of a group.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/group_list_members',
					common.baseUrl + '/myOrganization/groups/{Group.id}/members',
					false,				
					function () {
						doSnippet(groups.getMembers);
					}),
				new Snippet(
					'GET myOrganization/groups/{Group.id}/owners',
					'Gets the owners of a group.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/group_list_owners',
					common.baseUrl + '/myOrganization/groups/{Group.id}/owners',	
					false,			
					function () {
						doSnippet(groups.getOwners);
					})
			]
		};
		
		////////////////////////////////////////////////
		// All of the snippets that fall under the    //
		// 'drives' tenant-level resource collection. //
		////////////////////////////////////////////////
		var drivesSnippets = {
			groupTitle: 'drives',
			snippets: [
				/////////////////////////////////
				//       DRIVES SNIPPETS       // 
				/////////////////////////////////				
				new Snippet(
					'GET myOrganization/drives',
					'Gets  all of the drives in your tenant.',
					'http://graph.microsoft.io/docs/api-reference/v1.0/api/drive_get',
					common.baseUrl + '/myOrganization/drives',	
					false,				
					function () {
						doSnippet(drives.getDrives);
					})
			]
		};
		
		// Properties
		vm.activeSnippet;
		vm.snippetGroups = [
			usersSnippets,
			groupsSnippets,
			drivesSnippets
		];
		 
		// Methods
		vm.setActive = setActive;
		
		/////////////////////////////////////////
		// End of exposed properties and methods.
		
		var tenant;
		
		/**
		 * This function does any initialization work the 
		 * controller needs.
		 */
		(function activate() {
			if (adalAuthenticationService.userInfo.isAuthenticated) {
				vm.activeSnippet = vm.snippetGroups[0].snippets[0];
				
				tenant = adalAuthenticationService.userInfo.userName.split('@')[1];
			}
		})();
		
		/**
		 * Takes in a snippet, starts animation, executes snippet and handles response,
		 * then stops animation. 
		 */
		function doSnippet(snippet) {
			// Starts button animation.
			$scope.laddaLoading = true;
			
			// Clear old data.
			vm.activeSnippet.request = null;
			vm.activeSnippet.response = null;
			vm.activeSnippet.setupError = null;

			snippet()
				.then(function (response) {
					// Format request and response.
					var request = response.config;
					response = {
						status: response.status,
						statusText: response.statusText,
						data: response.data
					};

					// Attach response to view model.
					vm.activeSnippet.request = request;
					vm.activeSnippet.response = response;
				}, function (error) {
					// If a snippet requires setup (i.e. creating an event to update, creating a file
					// to delete, etc.) and it fails, handle that error message differently. 
					if (error.setupError) {
						// Extract setup error message.
						vm.activeSnippet.setupError = error.setupError;
						
						// Extract response data.
						vm.activeSnippet.response = {
							status: error.response.data,
							statusText: error.response.statusText,
							data: error.response.data
						}
						
						return;
					}
					
					// Format request and response.
					var request = error.config;
					error = {
						status: error.status,
						statusText: error.statusText,
						data: error.data
					};
					
					// Attach response to view model.
					vm.activeSnippet.request = request;
					vm.activeSnippet.response = error;
				})
				.finally(function () {
					// Stops button animation.
					$scope.laddaLoading = false;
				});
		};
		
		/**
		 * Sets class of list item in the sidebar. 
		 */
		function setActive(title) {
			if (!adalAuthenticationService.userInfo.isAuthenticated) {
				return;
			}

			if (title === vm.activeSnippet.title) {
				return 'active';
			}
			else {
				return '';
			}
		};

		/**
		 * Function to create partial functions. Taken
		 * from Stack Overflow user Jason Bunting.
		 * 
		 * http://stackoverflow.com/users/1790/jason-bunting
		 * http://stackoverflow.com/questions/321113/how-can-i-pre-set-arguments-in-javascript-
		 * function-call-partial-function-appli/321527#321527
		 */
		function partial(func /*, 0..n args */) {
			var args = Array.prototype.slice.call(arguments, 1);
			return function () {
				var allArguments = args.concat(Array.prototype.slice.call(arguments));
				return func.apply(this, allArguments);
			};
		};
	};
})();
