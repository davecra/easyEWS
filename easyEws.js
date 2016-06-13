/*!
 * easyEWS JavaScript Library v1.0.2
 * http://davecra.com
 *
 * Copyright David E. Craig and other contributors
 * Released under the MIT license
 * https://tldrlegal.com/license/mit-license
 *
 * Date: 2016-04-18T19:14EST
 */

var easyEws = (function () {
    "use strict";

    var easyEws = {};

    // PRIVATE: creates a SOAP EWS wrapper
    function getSoapHeader(request) {
        var result =
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '   <soap:Header>' +
            '       <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
            '   </soap:Header>' +
            '   <soap:Body>' + request + '</soap:Body>' +
            '</soap:Envelope>';
        return result;
    };

    // PRIVATE: makes an EWS callback with promise
    function asyncEws(soap, successCallback, errorCallback) {
        Office.context.mailbox.makeEwsRequestAsync(soap, function (ewsResult) {
            if (ewsResult.status == "succeeded") {
                var xmlDoc = $.parseXML(ewsResult.value);
                successCallback(xmlDoc); 
            } else {
                if (errorCallback != null)
                    errorCallback(ewsResult);
            }
        });
    };

    // PUBLIC: updates the x-headers in the mail item
    // RETUNS: 'succeeded' if call completed successfully
    // SEE: https://msdn.microsoft.com/en-us/library/office/dn596091(v=exchg.150).aspx
    easyEws.updateEwsHeader = function (mailItemId, headerName, headerValue, successCallback, errorCallback) {
        var soap =
            '<m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite">' +
            '   <m:ItemChanges>' +
            '       <t:ItemChange>' +
            '           <t:ItemId Id="' + mailItemId + '"/>' +
            '           <t:Updates>' +
            '               <t:SetItemField>' +
            '                   <t:ExtendedFieldURI DistinguishedPropertySetId="InternetHeaders"' +
            '                                       PropertyName="' + headerName + '"' +
            '                                       PropertyType="String" />' +
            '                   <t:Message>' +
            '                       <t:ExtendedProperty>' +
            '                           <t:ExtendedFieldURI DistinguishedPropertySetId="InternetHeaders"' +
            '                                               PropertyName="' + headerName + '"' +
            '                                               PropertyType="String" />' +
            '                               <t:Value>' + headerValue + '</t:Value>' +
            '                           </t:ExtendedProperty>' +
            '                   </t:Message>' +
            '               </t:SetItemField>' +
            '           </t:Updates>' +
            '       </t:ItemChange>' +
            '   </m:ItemChanges>' +
            '</m:UpdateItem>';


        soap = getSoapHeader(soap);
        // make the EWS call
        asyncEws(soap, function (xmlDoc) {
            successCallback("succeeded");
        }, function (errorDetails) {
            if (errorCallback != null)
                errorCallback(errorDetails);
        });
    };

    // PUBLIC:  returns a list of items in the folder
    // RETURNS: an array of ItemIds
    easyEws.getFolderItemIds = function (folderId, successCallback, errorCallback) {
        var soap =
            '<m:FindItem Traversal="Shallow">' +
            '   <m:ItemShape> ' +
            '       <t:BaseShape>IdOnly</t:BaseShape>' +
            '   </m:ItemShape>' +
            '   <m:ParentFolderIds>' +
            '       <t:FolderId Id="' + folderId + '"/>' +
            '   </m:ParentFolderIds>' +
            '</m:FindItem>';

        var returnArray = [];
        soap = getSoapHeader(soap);

        // call ews
        asyncEws(soap, function (xmlDoc) {
            $.each(xmlDoc.getElementsByTagName("t:ItemId"), function (index, value) {
                returnArray.push(value.getAttribute("Id"));
            });
            successCallback(returnArray);
        }, function (errorDetails) {
            if (errorCallback != null) {
                errorCallback(errorDetails);
            }
        });
    }

    // PUBLIC:  gets the details for a specific item by ID
    // RETURNS: a Dictionary of key/value pairs for the mail item
    easyEws.getMailItem = function(ItemId, successCallback, errorCallback) {
        var soap =
            '<m:GetItem>' +
            '   <m:ItemShape>' +
            '       <t:BaseShape>Default</t:BaseShape>' +
            '       <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
            '   </m:ItemShape>' +
            '   <m:ItemIds>' +
            '       <t:ItemId Id="' + ItemId + '" />' +
            '   </m:ItemIds>' +
            '</m:GetItem>';
        soap = getSoapHeader(soap);
        // make call to EWS
        asyncEws(soap, function (xmlDoc) {
            var item = new MailItem(xmlDoc);
            successCallback(item);
        }, function (errorDetails) {
            if(errorCallback != null) {
                errorCallback(errorDetails);
            }
        });
    }

    // PUBLIC:  expand a group and returns all the members
    // NOTE:    does not enumerate groups in groups
    // RETURNS: An array of Email Addresses
    easyEws.expandGroup = function (group, successCallback, errorCallback) {
        var soap =
            '<m:ExpandDL>' +
            '    <m:Mailbox>' +
            '        <t:EmailAddress>" + group + "</t:EmailAddress>' +
            '    </m:Mailbox>' +
            '</m:ExpandDL>';
        soap = getSoapHeader(soap);
        // make the EWS call
        var returnArray = [];
        asyncEws(soap, function (xmlDoc) {
            var extendedProps = xmlDoc.getElementsByTagName("EmailAddress");
            $.each(extendedProps, function (index, value) {
                returnArray.push(value);
            });
            successCallback(returnArray);
        }, function (errorDetails) {
            if (errorCallback != null)
                errorCallback(errorDetails);
        });
    };

    // PUBLIC: Find a given conversation by the ID
    // RETURNS: An array of ItemID
    easyEws.findConversationItems = function (conversationId, successCallback, errorCallback) {
        // NOTE: search for parent:
        // http://stackoverflow.com/questions/19008696/exchange-find-items-in-ews-conversation-using-xml-request
        // http://www.outlookcode.com/codedetail.aspx?id=1714
        // https://msdn.microsoft.com/en-us/library/office/dn610351(v=exchg.150).aspx
        var soap =
            '       <m:GetConversationItems>' +
            '           <m:ItemShape>' +
            '               <t:BaseShape>IdOnly</t:BaseShape>' +
            '               <t:AdditionalProperties>' +
            '                   <t:FieldURI FieldURI="item:Subject" />' +
            '                   <t:FieldURI FieldURI="item:DateTimeReceived" />' +
            '               </t:AdditionalProperties>' +
            '           </m:ItemShape>' +
            '           <m:FoldersToIgnore>' +
            '               <t:DistinguishedFolderId Id="deleteditems" />' +
            '               <t:DistinguishedFolderId Id="drafts" />' +
            '           </m:FoldersToIgnore>' +
            '           <m:SortOrder>TreeOrderDescending</m:SortOrder>' +
            '           <m:Conversations>' +
            '               <t:Conversation>' +
            '                   <t:ConversationId Id="' + conversationId + '" />' +
            '               </t:Conversation>' +
            '           </m:Conversations>' +
            '       </m:GetConversationItems>';
        soap = getSoapHeader(soap);
        // Make EWS call
        asyncEws(soap, function (xmlDoc) {
            var returnArray = [];
            $.each(xmlDoc.getElementsByTagName("t:ItemId"), function (index, value) {
                returnArray.push(value.getAttribute("Id"));
            });
            successCallback(returnArray);
        }, function (errorDetails) {
            if (errorCallback != null)
                errorCallback(errorDetails);
        });
    };

    // PUBLIC Gets Internet headers for a spific item
    // RETURNS: a Dictionary of key value pairs
    // SEE: https://msdn.microsoft.com/en-us/library/office/aa566013(v=exchg.150).aspx
    easyEws.getEwsHeaders = function (itemId, successCallback, errorCallback) {
        var soap =
        '   <m:GetItem>' +
        '       <m:ItemShape>' +
        '           <t:BaseShape>AllProperties</t:BaseShape>' +
        '           <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
        '       </m:ItemShape>' +
        '       <m:ItemIds>' +
        '           <t:ItemId Id="' + itemId + '" />' +
        '       </m:ItemIds>' +
        '   </m:GetItem>';

        soap = getSoapHeader(soap);
        // Make the EWS call
        var returnArray = new Dictionary(); 
        asyncEws(soap, function (xmlDoc) {
            for (var item in xmlDoc.getElementsByTagName("t:InternetMessageHeader")) {
                returnArray.add(item.getAttribute("HeaderName"), item.textContent);
            }
            successCallback(returnArray);
        }, function (errorDetails) {
            if (errorCallback != null)
                errorCallback(errorDetails);
        });
    };

    // PUBLIC: updates a folder property
    // RETURNS: 'succeeded' is process completed successfully
    easyEws.updateFolderProperty = function (folderId, propName, propValue, successCallback, errorCallback) {
        var soap =
            '       <m:UpdateFolder>' +
            '           <m:FolderChanges>' +
            '               <t:FolderChange>' +
            '                   <t:FolderId Id="' + folderId + '" />' +
            '                   <t:Updates>' +
            '                       <t:SetFolderField>' +
            '                           <t:ExtendedFieldURI ' +
            '                              DistinguishedPropertySetId="PublicStrings" ' +
            '                              PropertyName="' + propName + '" ' +
            '                              PropertyType="String" />' +
            '                            <t:Folder>' +
            '                               <t:ExtendedProperty>' +
            '                                  <t:ExtendedFieldURI ' +
            '                                     DistinguishedPropertySetId="PublicStrings" ' +
            '                                     PropertyName="' + propName + '" ' +
            '                                     PropertyType="String" />' +
            '                                 <t:Value>' + propValue + '</t:Value>' +
            '                              </t:ExtendedProperty>' +
            '                           </t:Folder>' +
            '                       </t:SetFolderField>' +
            '                   </t:Updates>' +
            '               </t:FolderChange>' +
            '           </m:FolderChanges>' +
            '       </m:UpdateFolder>';

        soap = getSoapHeader(soap);
        // make the EWS call
        asyncEws(soap, successCallback, errorCallback);
    }

    // PUBLIC: gets a folder property
    // RETURNS: property value if process completed successfully
    easyEws.getFolderProperty = function (folderId, propName, successCallback, errorCallback) {

        var soap =
            '<m:GetFolder>' +
                '<m:FolderShape>' +
                    '<t:BaseShape>IdOnly</t:BaseShape>' +
                    '<t:AdditionalProperties>' +
                        '<t:ExtendedFieldURI ' +
                        '   DistinguishedPropertySetId="PublicStrings" ' +
                        '   PropertyName="' + propName + '" ' +
                        '   PropertyType="String" />' +
                    '</t:AdditionalProperties>' +
                '</m:FolderShape>' +
                '<m:FolderIds>' +
                    '<t:FolderId Id="' + folderId + '"/>' +
                '</m:FolderIds>' +
            '</m:GetFolder>';
        soap = getSoapHeader(soap);
        // make the EWS call
        asyncEws(soap, function(xmlDoc) {
            successCallback(xmlDoc.getElementsByTagName("t:Value")[0].textContent);
        }, errorCallback);
    }

    // PUBLIC: Gets the folder id by the given name from the store
    // RETURNS: a string with ID of the folder
    easyEws.getFolderId = function (folderName, successCallback, errorCallback) {
        var soap =
            '    <m:GetFolder>' +
            '      <m:FolderShape>' +
            '        <t:BaseShape>IdOnly</t:BaseShape>' +
            '      </m:FolderShape>' +
            '      <m:FolderIds>' +
            '        <t:DistinguishedFolderId Id="' + folderName + '" />' +
            '      </m:FolderIds>' +
            '    </m:GetFolder>';
        soap = getSoapHeader(soap);
        // make EWS callback
        asyncEws(soap, function (xmlDoc) {
            var id = xmlDoc.getElementsByTagName("t:FolderId")[0].getAttribute("Id");
            successCallback(id);
        }, function (errorDetails) {
            if (errorCallback != null)
                errorCallback(errorDetails);
        });
    }
    
    return easyEws;
    
})();

/* HELPER FUNCTIONS AND CLASSES */
function MailItem(value) {

    this.value = value || {};

    MailItem.prototype.MimeContent = function () {
        return this.value.getElementsByTagName("t:MimeContent")[0].textContent;
    };

    MailItem.prototype.MimeContent.CharacterSet = function () {
        return this.value.getElementsByTagName("t:MimeContent")[0].getAttribute("CharacterSet");
    };

    MailItem.prototype.Subject = function () {
        return this.value.getElementsByTagName("t:Subject")[0].textContent;
    };
}

function Dictionary(values) {
    this.values = values || {};

    var forEachIn = function (object, action) {
        for (var property in object) {
            if (Object.prototype.hasOwnProperty.call(object, property))
                action(property, object[property]);
        }
    };

    Dictionary.prototype.containsKey = function (key) {
        return Object.prototype.hasOwnProperty.call(this.values, key) &&
          Object.prototype.propertyIsEnumerable.call(this.values, key);
    };

    Dictionary.prototype.forEach = function (action) {
        forEachIn(this.values, action);
    };

    Dictionary.prototype.lookup = function (key) {
        return this.values[key];
    };

    Dictionary.prototype.add = function (key, value) {
        this.values[key] = value;
    };

    Dictionary.prototype.length = function () {
        var len = 0;
        forEachIn(this.values, function () { len++ });
        return len;
    };
};
