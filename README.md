![LOGO](https://davecra.files.wordpress.com/2017/07/easyews.png?w=600)
# Introduction
This library makes performing EWS operations from Outlook Mail Web Add-ins via JavaScript much easier. EWS is quite difficult to perform from JavaScript because you have to format a specific SOAP message in order to submit the request using [makeEwsRequestAsync()](https://docs.microsoft.com/en-us/outlook/add-ins/web-services?product=outlook). However, this is complicated by the fact you then get a SOAP message back that you then have to parse in order to get your result (or error). This library limits your need to call makeEwsRequestAsync() by encapsulating the call in easy to use functions.

**NOTE:** Microsoft official guidance at this point is to no longer use EWS, but rather to use the REST API's. Some of this functionality (as of this writing: 8/1/2017), is available through REST and some is not. However, to get more informaiton, please see the following link:https://docs.microsoft.com/en-us/outlook/add-ins/use-rest-api

### Installation
To install this library, run the following command:

```
npm -install easyews
```

### Follow
Please follow my blog for the latest developments on easyEws. You can find my blog here:

![LOGO](https://davecra.files.wordpress.com/2017/07/blog-icon-large.png?w=20) http://theofficecontext.com

You can use this link to narrow the results only to those posts which relate to this library:

* https://theofficecontext.com/?s=easyews

![TWITTER](https://davecra.files.wordpress.com/2010/10/tlogo.png?w=20) You can also follow me on Twitter: [@davecra](http://twitter.com/davecra)

![LINKEDIN](https://davecra.files.wordpress.com/2014/02/inbug-60px-r.png?w=20) And also on LinkedIn: [davidcr](https://www.linkedin.com/in/davidcr/)

# Usage
This section is covers how to use easyEws. The following functions are available to call:

* [sendPlainTextEmailWithAttachment](#sendPlainTextEmailWithAttachment) - creates a new emails message with a single attachment and sends it
* [getMailItemMimeContent](#getMailItemMimeContent)- gets the mail item as raw MIME data
* [updateEwsHeader](#updateEwsHeader) - Updates the headers in the mail item
* [getFolderItemIds](#getFolderItemIds)- Returns a list of items in the folder
* [getMailItem](#getMailItem) - Gets the item details for a specific item by ID
* DO NOT USE! --> [expandGroup](#expandGroup) - The ExpandDL method is not supported
* [findConversationItems](#findConversationItems) - Find a given conversation by the ID
* [getSpecificHeader](#getSpecificHeader) - Gets a specific Internet header for a spific item
* [getEwsHeaders](#getEwsHeaders) - Gets Internet headers for a spific item
* [updateFolderProperty](#updateFolderProperty) - Updates a folder property. If the property does not exist, it will be created.
* [getFolderProperty](#getFolderProperty) -  Gets a folder property
* [getFolderId](#getFolderId) - Gets the folder id by the given name from the store
* [moveItem](#moveItem) - Moves an item from one folder to another

### sendPlainTextEmailWithAttachment <a name="sendPlainTextEmailWithAttachment"></a>
This method will send a plain text message to a recipient with an attachment. This function is very specific, but provides the essential foundation for creating an email with different options.

**NOTE**: If additional options are needed, different types of send requests, please contact me.

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
Example is TBD.
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
Example is TBD.
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
Example is TBD.
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
**DO NOT USE**. This does not function in the current makeEwsRequestAsync() interface. As such, you cannot use this method. This is currently in the list for future possible use.

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
Example is TBD.
```

### findConversationItems <a name="findConversationItems"></a>
This method will return all the related itemID's in a specific conversation. If you need to find a specific item you can then use the ID's to make the [getMailItem()](#getMailItem) method.

Here are the paramaters for this method:
* **conversationId**: *string* - the conversation ID for which you want to retrieve all the related items
* successCallback
* **errorCallback**: *function(**error**: string)* - If an error occurs a string with the resulting error will be returned. For more detail on the exact nature of the issue, you can refer to the debugCallback.
56
* **debugCallback**: *function(**debug**: string)* - Contains a detailed XML output with the original xml sent, the response from the server in xml, and any status messages or error objects returned. 

##### Example #####
Here is an example of how to use this method:

```javascript
Example is TBD.
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
Example is TBD.
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
Example is TBD.
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
Example is TBD.
```