![LOGO](https://davecra.files.wordpress.com/2017/07/easyews.png?w=600)
# Introduction
This library makes performing EWS operations from Outlook Mail Web Add-ins via JavaScript much easier. EWS is quite difficult to perform from JavaScript because you have to format a specific SOAP message in order to submit the request using [makeEwsRequestAsync()](https://docs.microsoft.com/en-us/outlook/add-ins/web-services?product=outlook). However, this is complicated by the fact you then get a SOAP message back that you then have to parse in order to get your result (or error). This library limits your need to call makeEwsRequestAsync() by encapsulating the call in easy to use functions.

**NOTE**: If you encounter any problems with this library, please submit an issue: https://github.com/davecra/easyEWS/issues.

**NOTE:** Microsoft official guidance at this point is to no longer use EWS, but rather to use the REST API's. Some of this functionality (as of this writing: 8/1/2017), is available through REST and some is not. However, to get more informaiton, please see the following link:https://docs.microsoft.com/en-us/outlook/add-ins/use-rest-api

### Installation
To install this library, run the following command:

```
npm -install easyews
```

### Referencing
easyEws comes with both a full (debug) version and a minified version. To access the debug version from node_modules (if your source is in the root of your project):

```html
<script type="text/javascript" src="node_modules/easyews/easyews.js"></script>
```

To access the minified version from node_modules (if your source is in the root of your project):

```html
<script type="text/javascript" src="node_modules/easyews/easyews.min.js"></script>
```

easyEws can also be accessed from the following CDN: https://cdn.jsdelivr.net/npm/easyews/easyEws.js

```html
<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/easyews/easyEws.js"></script>
```

or, the minified version:

```html
<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/easyews/easyEws.min.js"></script>
```

**NOTE**: This uses the CDN published through NPM. In the past this pointed to the one on GitHub, but I have had versioning issues where JSDelivr does not update for serveral days or at all with GitHub updates. However, the NPM CDN pointer seems to work much better.

**NOTE**: If you need to reference a spcific version of EWS, every version is kept and accessible via version number. The following table contains the most recent versions:

| Version | Url                                                     | Minified                                                     |
|:-------:|:-------------------------------------------------------:|:------------------------------------------------------------:|
|v1.0.14  | https://cdn.jsdelivr.net/npm/easyews@1.0.14/easyEws.js  | https://cdn.jsdelivr.net/npm/easyews@1.0.14/easyEws.min.js   |
|v1.0.15  | https://cdn.jsdelivr.net/npm/easyews@1.0.15/easyEws.js  | https://cdn.jsdelivr.net/npm/easyews@1.0.15/easyEws.min.js   |
|v1.0.16  | https://cdn.jsdelivr.net/npm/easyews@1.0.16/easyEws.js  | https://cdn.jsdelivr.net/npm/easyews@1.0.16/easyEws.min.js   |
|v1.0.17  | https://cdn.jsdelivr.net/npm/easyews@1.0.17/easyEws.js  | https://cdn.jsdelivr.net/npm/easyews@1.0.17/easyEws.min.js   |
|v1.0.18  | https://cdn.jsdelivr.net/npm/easyews@1.0.18/easyEws.js  | https://cdn.jsdelivr.net/npm/easyews@1.0.18/easyEws.min.js   |
|v1.0.19  | https://cdn.jsdelivr.net/npm/easyews@1.0.19/easyEws.js  | https://cdn.jsdelivr.net/npm/easyews@1.0.19/easyEws.min.js   |
|v1.0.20  | https://cdn.jsdelivr.net/npm/easyews@1.0.20/easyEws.js  | https://cdn.jsdelivr.net/npm/easyews@1.0.20/easyEws.min.js   |

### Follow
Please follow my blog for the latest developments on easyEws. You can find my blog here:

![LOGO](https://davecra.files.wordpress.com/2017/07/blog-icon-large.png?w=20) http://theofficecontext.com

You can use this link to narrow the results only to those posts which relate to this library:

* https://theofficecontext.com/?s=easyews

![TWITTER](https://davecra.files.wordpress.com/2010/10/tlogo.png?w=20) You can also follow me on Twitter: [@davecra](http://twitter.com/davecra)

![LINKEDIN](https://davecra.files.wordpress.com/2014/02/inbug-60px-r.png?w=20) And also on LinkedIn: [davidcr](https://www.linkedin.com/in/davidcr/)

# Usage
This section is covers how to use easyEws. The following functions are available to call:

* [sendMailItem](#sendMailItem) - creates an email to multiple recipients, with or without attachments and sends it
* [sendPlainTextEmailWithAttachment](#sendPlainTextEmailWithAttachment) - creates a new emails message with a single mail item attachment (mime) and sends it
* [getMailItemMimeContent](#getMailItemMimeContent)- gets the mail item as raw MIME data
* [updateEwsHeader](#updateEwsHeader) - Updates the headers in the mail item
* [getFolderItemIds](#getFolderItemIds)- Returns a list of items in the folder
* [getMailItem](#getMailItem) - Gets the item details for a specific item by ID
* [expandGroup](#expandGroup) - Returns a list of members to an Exchange Distribution Group
* [splitGroupsAsync](#splitGroupsAsync) - Returns a list of all users found in every group (and groups in groups, etc.)
* [getAllRecipientsAsync](#getAllRecipientsAsync) - Returns lists of users and groups on the To/CC/BCC
* [findConversationItems](#findConversationItems) - Find a given conversation by the ID
* [getSpecificHeader](#getSpecificHeader) - Gets a specific Internet header for a spific item
* [getEwsHeaders](#getEwsHeaders) - Gets Internet headers for a spific item
* [updateFolderProperty](#updateFolderProperty) - Updates a folder property. If the property does not exist, it will be created.
* [getFolderProperty](#getFolderProperty) -  Gets a folder property
* [getFolderId](#getFolderId) - Gets the folder id by the given name from the store
* [moveItem](#moveItem) - Moves an item from one folder to another
* [resolveRecipient](#resolveRecipient) - Resolves a recipient
* [getParentId](#getParentId) - Gets the Id for the parent of the specified mail item

### sendMailItem <a name="sendMailItem"></a>
This is a ALL PURPOSE method to send an HTML or Text message to multiple recipients on the TO line with any number of attachments of either Mail Item or File types. 

Here are the paramaters for this method:
* **p**: **SendMailFunctionObject** - this is an object that defines the message you want to send.

Here are the parameters for the SendMailFunctionObject: 
* **subject**: *string* - this is the subject for the email to be set
* **body**: *string* - this is the body of the message to be sent. It can be plain text or HTML. You also do not need to escape your HTML, it will be escaped for you.
* **bodytype**: *string* - this is either "html" or "text". The default is "text"
* **recipients**: *string[]* - this is an array of email addresses
* **attachments**: *SimpleAttachmentObject[]* - array of simple attachment objects. Pass [{}] if no attachments. Here is the SimpleAttachmentObject definition:
	* **SimpleAttachmentObject.name**: *string* - the name of the attachment
	* **SimpleAttachmentObject.mime**: *string* - (base64 string) mime content for the attachment
	* **SimpleAttachmentObject.type**: *string* - this is either "file" (for any type of file) or "item" (for a mail item). Default is "file"
* **folderid**: *string* - distinguished folder id of folder to put the sent item in
* **successCallback**: *function(**result**: string)* - Returns "success" if completed successfully.
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returne

##### Example #####
Here is an example of how to use this method

```javascript
  /**@type {SimpleAttachmentObject} */
  var att = new SimpleAttachmentObject("welcome_packet.txt", "SGVsbG8gd29ybGQh", "file");
  /**@type {SendMailFunctionObject} */
  var p = new SendMailFunctionObject("Simple Subject", 
                                     "<b>Welcome</b> and hello World!", 
                                     "html", 
                                      ["testing@contoso.com"], 
                                      [att], "sentitems", 
                                      function(result) {
                                        console.log(result);
                                      }, function(error) {
                                        console.log(error);
                                      }, function(debug) {
                                        console.log(debug);
                                      });
  easyEws.sendMailItem(p);
```

### sendPlainTextEmailWithAttachment <a name="sendPlainTextEmailWithAttachment"></a>
This method will send a plain text message to a recipient with a mime message item attachment. This function is very specific, but provides the essential foundation for creating an email with different options.

Here are the paramaters for this method:
* **subject**: *string* - this is the subject for the email to be set
* **body**: *string* - this is the body of the message to be sent. It must be in plain text. HTML is NOT supported. 
* **to**: *string* - this is the list of recipients for the email to be sent
* **attachmentName**: *string* - this is the name of a single attachment you can apply to the email message
* **attachmentMime**: *string* - this is the MIME content (base64) of the object to be attached.
* **successCallback**: *function(**result**: string)* - Returns "success" if completed successfully.
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
function sendSuspiciousMessage() {
    var item = Office.context.mailbox.item;
    var itemId = item.itemId;
    easyEws.getMailItemMimeContent(itemId, function(mimeContent) {
    	var toAddress = "davidcr@outlook.com";
    	easyEws.sendPlainTextEmailWithAttachment("Suspicious Email Alert",
                                                 "A user has forwarded a suspicious email",
                                                 toAddress,
                                                 "Suspicious_Email.eml",
                                                 mimeContent,
                                                function(result) { console.log(result); },
                                                function(error) { console.log(error); }, 
                                                function (debug) { console.log(debug); });
  }, function(error) { console.log(error); }, function(debug) { console.log(debug); });
}
```

### getMailItemMimeContent <a name="getMailItemMimeContent"></a>
This method will return the MIME content of a specific mail item.

Here are the paramters for this method:
* **mailItemId**: *string* - the mail itemID for which you want to retrieve MIME content
* successCallback  - function(result: string) - Returns the MIME content as a BASE64 string. You will want to btoa() the results to manage it as an object or string
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
function sendSuspiciousMessage() {
	var item = Office.context.mailbox.item;
	itemId = item.itemId;
	mailbox = Office.context.mailbox;
	easyEws.getMailItemMimeContent(itemId, function(mimeContent) {
		var toAddress = "securityteam@somwhere.local";
		easyEws.sendPlainTextEmailWithAttachment("Suspicious Email Alert",
							 "A user has forwarded a suspicious email",
							 toAddress,
							 "Suspicious_Email.eml",
							 mimeContent,
							 function(result) { console.log(result); },
							 function(error) { console.log(error); }, 
							 function (debug) { console.log(debug); }
		);
	 }, function(error) { console.log(error); },
	    function(debug) { console.log(debug); } 
	);
}
```

### updateEwsHeader <a name="updateEwsHeader"></a>
This method will update the mail transport (x-headers) in the selected message. The message must be saved into the mail store for this to work.

**NOTE**: If you try to perform this on Outlook 2016 / full-windows-client, the settings may not stick if you are running in cached mode. This is by default. The only way for this to work is to wait for 30 seconds to a minute. This problem DOES NOT occur with Outlook in online mode and in Outlook Web Access (OWA) or Outlook Online (Office 365).

Here are the paramters for this method:
* **mailItemId**: *string* - The mail item id for the items whos header you want to update
* **headerName**: *string* - The header you want to update
* **headerValue**: *string* - The value you want to set the header to
* **successCallback**: *function(**result**: string)* - Returns "success" if completed successfully.
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
// first, we have to save the item to the drafts folder, which will get us the 
// Exchange EWS_ID of the item. We then use that ID to update the header.
Office.context.mailbox.item.saveAsync(function (idResult) {
    	var id = idResult.value;
	// now that we have the ID of the mail item, we update the header
	easyEws.updateEwsHeader(id, "x-myheader", value, false, function () {
	    console.log("x-header has been set.");
	}, function(error) {
	    console.log(error);
	}, function(debug) {
	    console.log(debug);
	});
});
```

### getFolderItemIds <a name="getFolderItemIds"></a>
This method will return an array of item IDs for all the items found in a particular folder. If you need the number of items in a folder you can get the array count. If you need to find a specific item you can then use the ID's to make the [getMailItem()](#getMailItem) method.

Here are the paramaters for this method:
* **folderId**: *string*: The folder ID for which you want to get all the items
* **successCallback**: *function(**result**: string[])* - returns an array of itemID's if completed successfully.
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
Example is TBD.
```

### getMailItem <a name="getMailItem"></a>
This method will get all the details of a mail item and store it in a MailItem object.

**NOTE:** The MailItem object right now is very primitive only surfacing a few of the properties. If more proeprties are needed, please contact me.

Here are the parameters for this method:
* **itemId**: *string* - the item id you want to access.
* **successCallback**: *function(**result**: MailItem)* - Returns a MailItem object with the values of the email
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
Example is TBD.
```

### expandGroup <a name="expandGroup"></a>
(2/27/2018) This does not function in Exchange 2016 / On-Prem (TBD in CU9). 

This method will take a group name and split it one level to constituent users and groups. It is not recursive, so if you need to split multiple groups within groups, you will need to call this function multiple times.

Here are the paramters for this method:
* **group**: *string* - the name of the group you want to expand. 
* **successCallback**: *function(**result**: MailBoxUser[])* - If successful will return an array of MailBoxUser objects.
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
/**
 * Breaks upp all the groups on the to/cc/bcc lines of the message
 * and reports the total number of member affected.
 * @param {Office.AddinCommands.Event} event 
 */
function breakGroups(event) {
  $(document).ready(function () {
    var mailItem = Office.context.mailbox.item;
    // Get all the recipients for the composed mail item
    easyEws.getAllRecipientsAsync(mailItem, 
    /**
     * Returns a groups of users[] and groups[] we only
     * care about the groups here and will expand them
     * @param {Office.EmailAddressDetails[]} users 
     * @param {Office.EmailAddressDetails[]} groups 
     */
    function(users,groups) {
      var allMembers = [];
      // loop through all the groups found
      for(var i=0;i<groups.length;i++) {
        var groupEmail = groups[i].emailAddress;
        easyEws.expandGroup(groupEmail, function(members) {
          console.log("Split group " + groupEmail + " with " + members.length + " member(s)");
          $.each(members, function(index, item) { allMembers.push(item); });
          // on the last item, notify the user
          if(i >= groups.length) {
            // NOTE: using OfficeJS.dialogs (https://github.com/davecra/officejs.dialogs)
            Alert.Show("All groups have been split.\nThere are " + allMembers.length + " members.", 
            function() {
              // button event complete
              event.completed();
            });
          }
        }, function(error) {
          console.log("ERROR: " + error.description);
          event.completed();
        }, function(debug) {
          // DEBUG OUTPUT:
          // displays the SOAP to EWS
          // displays the SOAP back from EWS
          console.log(debug);
        });
      }
    });
  });
}
```

### splitGroupsAsync <a name="splitGroupsAsync"></a>
This method takes an array of email addresses (as strings) for Distribution Groups and makes recursive asynchronous calls to [expandGroup](#expandGroup) to return a list of unique users found inside each group, and even groups within groups (up to 100 groups total).

NOTE: This function will automatically exit after 100 groups (including groups within groups) have been parsed. This is to prevent a possible hang when encountering circular groups.

Here are the paramters for this method:
* **groups**: *string[]* - an array of email addresses (or group names) for Distribution Lists you want to expand. 
* **successCallback**: *function(**result**: MailBoxUser[])* - If successful will return an array of all unique MailBoxUser objects found within all groups (and groups within groups).
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
/**
 * Breaks upp all the groups on the to/cc/bcc lines of the message
 * and reports the total number of member affected.
 * @param {Office.AddinCommands.Event} event 
 */
function breakGroups(event) {
  $(document).ready(function () {
    var mailItem = Office.context.mailbox.item;
    // Get all the recipients for the composed mail item
    easyEws.getAllRecipientsAsync(mailItem, 
    /**
     * Returns a groups of users[] and groups[] we only
     * care about the groups here and will expand them
     * @param {Office.EmailAddressDetails[]} users 
     * @param {Office.EmailAddressDetails[]} groups 
     */
    function(users,groups) {
      /** @type {string[]} */
      var groupList = [];
      $.each(groups, function(index, item) { groupList.push(item.emailAddress); })
      easyEws.splitGroupsAsync(groupList, 
      /**
       * Success callback from splitGroupAsync, returns an array
       * of all the users mailboxes found in an array
       * @param {MailboxUser[]} members 
       */
      function(members) {
        console.log("Split " + groups.length + " groups containing " + members.length + " member(s)");
        // NOTE: using OfficeJS.dialogs (https://github.com/davecra/officejs.dialogs)
        Alert.Show("All groups have been split.\nThere are " + members.length + " members.", 
        function() {
          // button event complete
          event.completed();
        });
      }, function(error) {
        console.log("ERROR: " + error.description);
        event.completed();
      }, function(debug) {
        // DEBUG OUTPUT:
        // displays the SOAP to EWS
        // displays the SOAP back from EWS
        console.log(debug);
      });
    });
  });
}
```

### getAllRecipientsAsync <a name="getAllRecipientsAsync"></a>
This method accepts the current ComposeItem and then parse the To/CC/BCC lines and returns a list of unique users (as an array of MailBoxUser) and a unique list of groups (as an array of MailBoxUser). This function is useful in conjunction with the [splitGroupsAsync()](#splitGroupsAsync) method.

Here are the parameters of the method:
* **composeItem**: *Office.Types.ItemCompose* - the item currently being composed in the mail client. Uses this object to access the To/CC/BCC lines async.
* **successCallback**: *function(**users**: MailBoxUser[], **groups**: MailBoxUser[])* - If successful will return a list of users (as an array of MailBoxUser) and a list of groups (as an array of MailBoxUser)
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
/**
 * Returns the total number of users on the to/cc/bcc lines
 * @param {Office.AddinCommands.Event} event 
 */
function getAllUsers(event) {
  $(document).ready(function () {
    var mailItem = Office.context.mailbox.item;
    // Get all the recipients for the composed mail item
    easyEws.getAllRecipientsAsync(mailItem, 
    /**
     * Returns a groups of users[] and groups[] we only
     * care about the groups here and will expand them
     * @param {Office.EmailAddressDetails[]} users 
     * @param {Office.EmailAddressDetails[]} groups 
     */
    function(users,groups) {
      Alert.Show("There are " + users.length + " user(s) and " + groups.length + " group(s) being directly address.", 
      function() {
        event.completed();
      });
    });
  });
}
```

### findConversationItems <a name="findConversationItems"></a>
This method will return all the related itemID's in a specific conversation. If you need to find a specific item you can then use the ID's to make the [getMailItem()](#getMailItem) method.

Here are the paramaters for this method:
* **conversationId**: *string* - the conversation ID for which you want to retrieve all the related items
* **successCallback**: *function(**items**: string[])* - returns an array of conversation ID's found
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
easyEws.findConversationItems(conversationId, function (itemArray) {
	if (itemArray === null || itemArray.length === 0) {
		console.log("No_coversation_items_found");
		return;
	}
	// we will grab the first item as the newest
	var mostRecentId = itemArray[0];
	console.log("Most recent conversation is: " + mostRecentId);
	}, function (error) {
		console.log(error);
	}, function (debug) {
		console.log(debug);
	});
```

### getSpecificHeader <a name="getSpecificHeader"></a>
This method will return a specific mail header (x-header) value for a specific mail item.

Here are the paramters for this method:
* **itemId**: *string* - the mail item you want to access
* **headerName**: *string* - the header property you want to access
* **headerType**: *string* - the header value type you want to acess. Supports: String or Integer
* **successCallback**: function(**result**: string)* - if successfull will return a the value of the header or NULL if it was not found
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
easyEws.getSpecificHeader(id, "x-myheader", "String", function (result) {
	if (result === null || result === "") {
		console.log("not_found");
	} else {
		console.log("Result: " + result);
	}
}, function (error) {
	console.log(error);
}, function (debug) {
	console.log(debug);
});
```

### getEwsHeaders <a name="getEwsHeaders"></a>
This method will get return all the Internet Message Headers (x-headers) for a given mail item.

Here are the parameters for this method:
* **itemId**: *string* - the ID of mail item you want to retrieve
* **successCallback**: *function(**result**: Dictionary<string, object>)* - If successful will return a Dictionary with all the Internet Message Headers in key/value pairs which can be iterated using .forEach().
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
easyEws.getEwsHeaders(id, function (headersDictionary) {
	var classificationResult = "";
	headersDictionary.forEach(function (key, value) {
		if (key.toLowerCase().startsWith("x-myheader")) {
			// success
			completeCallback(value);
			return;
		}
	});
	completeCallback("no_header_found");
}, function (error) {
	console.log("Failed to get EWS Headers.\n" + error);
}, function (debug) {
	console.log("getEwsHeaders: " + debug);
});
```

### updateFolderProperty <a name="updateFolderProperty"></a>
This method will update a specific named property on a MAPI folder object in the Exchange message store.

Here are the paramters for this method:
* **folderId**: *string* - The ID of the folder you want to update 
* **propName**: *string* - The property you want to add or update
* **propValue**: *string* - The value of the property you wan to add or update
* **successCallback**: *function(**result**: string)* - Returns "success" if completed successfully.
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
Example is TBD.
```

### getFolderProperty <a name="getFolderProperty"></a>
This method will get the value of a specific named property on an API folder object in the Exchange message store.

Here are the paramters for this method:
* **folderId**: *string* - The ID of the folder you want to access
* **propName**: *string* - The property you want to return
* **successCallback**: *function(**result**: string)* - Returns the property folder value if completed successfully.
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
Example is TBD.
```

### getFolderId<a name="getFolderId"></a>
Gets the folder ID for a specific names MAPI folder in the Excahnge mail store.

Here are the parameters for this method:
* **folderName**: *string* - The name of the folder you want to get the ID for.
* **successCallback**: *function(**result**: string)* - Returns the folder ID if completed successfully.
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
Example is TBD.
```

### moveItem<a name="moveItem"></a>
Moves an item to the specified MAPI folder in the Exchange store.

Here are the paramters for this method:
* **itemId**: *string* - The item ID for the message, appointment or meeting that is to be moved.
* **folderID**: *string* The folder name or the folder ID of the MAPI folder where you want to move the item to.
* **successCallback**: *function(**result**: string)* - the success callback. will return 'success' if the process completes successfully.
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
    // moves the selected item in a Read add-in to 
    // the deleted items folder...
    var item = Office.context.mailbox.item;
    var itemId = item.itemId;
    easyEws.moveItem(itemId, "deleteditems", function() {
      console.log("success");
    }, function(error) {
      console.log(error);
    });
```

### resolveRecipient<a name="resolveRecipient"></a>
Resolves a recipient.

Here are the parameters for this method:
* **recipient**: *string* - The recipient name or email
* **successCallback**: *function(**result**: ResolveNamesType[])* - the success callback. Will return an array of resolved names. The returned type is defined as:
         * @param {string} name
         * @param {string} emailAddress
         * @param {string} routingType
         * @param {string} mailboxType
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
async function run() {
    var msg = Office.context.mailbox.item;
    msg.to.getAsync(function(asyncResult) {
        if(asyncResult.status == Office.AsyncResultStatus.Failed) {
            $("#recipientResult").text(asyncResult.error);
        } else {
            /** @type {Office.EmailAddressDetails[]} */
            var recips = asyncResult.value;
            // are there any recipients at all
            if(recips == null || recips.length == 0) {
                $("#recipientResult").html("NO RECIPIENTS");
            } else {
                /** @type {string} */
                var info  = "<br/>DISPLAY NAME: " + recips[0].displayName + "<br/>" +
                            "EMAIL_ADDRESS: " + recips[0].emailAddress + "<br/>" +
                            "RECIPIENT_TYPE: " + recips[0].recipientType;
                easyEws.resolveRecipient(recips[0].emailAddress, function(result) {
                    if(result == null || result.length == 0) {
                        info += "<br/>UNRESOLVED</br>";
                    } else {
                        info += "<br/>RESOLVED: " + result[0].MailBoxType;
                        info += "<br/>RESOLVED EMAIL: " + result[0].EmailAddress;
                    }
                    // write tot he form
                    $("#recipientResult").html(info);
                }, function(error) {
                    $("#recipientResult").text(error);
                }, function(debug) {
                    $("#debugResult").text(debug)
                });
            }
        }
    });
}
```

### getParentId<a name="getParentId"></a>
Gets the Id for the parent of the specified mail item

Here are the parameters for this method:
* **childId**: *string* - The child message id
* **successCallback**: *function(**result**: string)* - the success callback. 
 - Will return a string with the Id if the parent was found.
 - Will return NULL if no parent was found.
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
Office.cast.item.toMessageCompose(Office.context.mailbox.item).saveAsync(function (idResult) {
  var childId = idResult.value;
  easyEws.getParentId(childId, function (parentId) {
      if(parentId !== null) {
        console.log("The parent ID is: " + parentId);
      } else {
        console.log("The item has no parent, or it could not be found.");
      }
    }, function (error) {
      console.log(error);
    }, function (debug) {
      console.log(debug);
  });
});
```
