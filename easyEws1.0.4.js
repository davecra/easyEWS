/*!
 * easyEWS JavaScript Library v1.0.4
 * http://davecra.com
 * https://raw.githubusercontent.com/davecra/easyEWS/master/easyEws/js/easyEws.js
 * 
 * Copyright David E. Craig and other contributors
 * Released under the MIT license
 * https://tldrlegal.com/license/mit-license
 *
 * Date: 2016-08-03T02:35EST
 */
var easyEws = (function () {
    "use strict";

    var easyEws = {};

    easyEws.initialize = function () {
        /// <summary>
        /// PRIVATE: creates a SOAP EWS wrapper
        /// </summary>
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

        function asyncEws(soap, successCallback, errorCallback, debugCallback) {
            /// <summary>
            /// PRIVATE: Makes an EWS callback with promise
            /// </summary>
            /// <param name="soap" type="String">XML Soap message</param>
            /// <param name="successCallback" type="Function">Success callback - function(DOMDocument) { }</param>
            /// <param name="errorCallback" type="Function">Error handler callback - function(Error) { }</param>
            /// <param name="debugCallback" type="Function">Debug callback - function(String) { }</param>
            Office.context.mailbox.makeEwsRequestAsync(soap, function (ewsResult) {
                if (ewsResult.status == "succeeded") {
                    var xmlDoc = $.parseXML(ewsResult.value); $.each()
                    successCallback(xmlDoc);
                    debugCallback(ewsResult.value); // return raw result
                } else {
                    if (errorCallback != null) {
                        errorCallback("makeEwsRequestAsync failed.");
                        debugCallback(ewsResult.value); // return raw result
                    }
                }
            });
        };

        easyEws.updateEwsHeader = function (mailItemId, headerName, headerValue,
                                            successCallback, errorCallback, debugCallback) {
            /// <summary>
            /// PUBLIC: Updates the x-headers in the mail item
            /// SEE: https://msdn.microsoft.com/en-us/library/office/dn596091(v=exchg.150).aspx
            /// </summary>
            /// <param name="mailItemId" type="String">The id of the item to update</param>
            /// <param name="headerName" type="String">The header item to add/update</param>
            /// <param name="headerValue" type="String">The header value to update</param>
            /// <param name="successCallback" type="Function">returns 'succeeeded' is successful - function(String) { }</param>
            /// <param name="errorCallback" type="Function">Error handler callback - function(Error) { }</param>
            /// <param name="debugCallback" type="Function">Debug handler returns raw XML - function(String) { }</param>
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
            }, function (debug) {
                if (debugCallback != null)
                    debugCallback(debug);
            });
        };

        easyEws.sendPlainTextEmailWithAttachment = function (subject, body, to, attachmentName, attachmentMime, successCallback, errorCallback) {
            /// <summary>
            /// PUBLIC: creates a new emails message with a single attachment and sends it
            /// </summary>
            /// <param name="subject" type="String">The subject for the message to be sent</param>
            /// <param name="body" type="String">The body of the message to be sent</param>
            /// <param name="to" type="String">The email address of the recipient</param>
            /// <param name="attachmentName" type="String">Name of the attachment</param>
            /// <param name="attachmentMime" type="String">MIME content in Base64 for the attachment</param>
            /// <param name="successCallback" type="Function">Callback with 'success' if compelted successfully - function(string) { }</param>
            /// <param name="errorCallback" type="Function">Error handler callback - function(Error) { }</param>
            /// <param name="debugCallback" type="Function">Debug handler returns raw XML - function(String) { }</param>
            var soap = '<m:CreateItem MessageDisposition="SendAndSaveCopy">' +
                       '    <m:Items>' +
                       '        <t:Message>' +
                       '            <t:Subject>' + subject + '</t:Subject>' +
                       '            <t:Body BodyType="Text">' + body + '</t:Body>' +
                       '            <t:Attachments>' +
                       '                <t:ItemAttachment>' +
                       '                    <t:Name>' + attachmentName + '</t:Name>' +
                       '                    <t:IsInline>false</t:IsInline>' +
                       '                    <t:Message>' +
                       '                        <t:MimeContent CharacterSet="UTF-8">' + attachmentMime + '</t:MimeContent>' +
                       '                    </t:Message>' +
                       '                </t:ItemAttachment>' +
                       '            </t:Attachments>' +
                       '            <t:ToRecipients><t:Mailbox><t:EmailAddress>' + to + '</t:EmailAddress></t:Mailbox></t:ToRecipients>' +
                       '        </t:Message>' +
                       '    </m:Items>' +
                       '</m:CreateItem>';

            soap = getSoapHeader(soap);

            // make the EWS call 
            asyncEws(soap, function (xmlDoc) {
                // Get the required response, and if it's NoError then all has succeeded, so tell the user.
                // Otherwise, tell them what the problem was. (E.G. Recipient email addresses might have been
                // entered incorrectly --- try it and see for yourself what happens!!)
                var result = xmlDoc.getElementsByTagName("ResponseCode")[0].textContent;
                if (result == "NoError") {
                    successCallback(result);
                }
                else {
                    if (errorCallback != null)
                        errorCallback(result);
                }
            }, function (errorDetails) {
                if (errorCallback != null)
                    errorCallback(errorDetails);
            }, function (debug) {
                if (debugCallback != null)
                    debugCallback(debug);
            });
        };

        easyEws.getMailItemMimeContent = function (mailItemId, successCallback, errorCallback, debugCallback) {
            /// <summary>
            /// PUBLIC: gets the mail item as raw MIME data
            /// </summary>
            /// <param name="mailItemId" type="type"></param>
            /// <param name="successCallback" type="Function">Callback with email message as MIME Base64 string - function(string) { } </param>
            /// <param name="errorCallback" type="Function">Error handler callback - function(Error) { }</param>
            /// <param name="debugCallback" type="Function">Debug handler returns raw XML - function(String) { }</param>
            var soap =
                '<m:GetItem>' +
                '    <m:ItemShape>' +
                '        <t:BaseShape>IdOnly</t:BaseShape>' +
                '        <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
                '    </m:ItemShape>' +
                '    <m:ItemIds>' +
                '        <t:ItemId Id="' + mailItemId + '"/>' +
                '    </m:ItemIds>' +
                '</m:GetItem>';
            soap = getSoapHeader(soap);
            // make the EWS call 
            asyncEws(soap, function (xmlDoc) {
                //var content = xmlDoc.getElementsByTagName("MimeContent")[0].textContent;
                successCallback(xmlDoc);
            }, function (errorDetails) {
                if (errorCallback != null)
                    errorCallback(errorDetails);
            }, function (debug) {
                if (debugCallback != null)
                    debugCallback(debug);
            });
        };


        easyEws.getFolderItemIds = function (folderId, successCallback, errorCallback, debugCallback) {
            /// <summary>
            /// PUBLIC:  Returns a list of items in the folder
            /// </summary>
            /// <param name="folderId" type="String">The ID of the folder you want to search</param>
            /// <param name="successCallback" type="Function">Callback with array of item IDs - function(String[]) { }</param>
            /// <param name="errorCallback" type="Function">Error handler callback - function(Error) { }</param>
            /// <param name="debugCallback" type="Function">Debug handler returns raw XML - function(String) { }</param>
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
            }, function (debug) {
                if (debugCallback != null)
                    debugCallback(debug);
            });
        }


        easyEws.getMailItem = function (ItemId, successCallback, errorCallback, debugCallback) {
            /// <summary>
            /// PUBLIC:  Gets the item details for a specific item by ID
            /// </summary>
            /// <param name="ItemId" type="String">The ID for the item</param>
            /// <param name="successCallback" type="Function">Callback with the details of the MailItem - function(MailItem) { }</param>
            /// <param name="errorCallback" type="Function">Error handler callback - function(Error) { }</param>
            /// <param name="debugCallback" type="Function">Debug handler returns raw XML - function(String) { }</param>
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
                if (errorCallback != null) {
                    errorCallback(errorDetails);
                }
            }, function (debug) {
                if (debugCallback != null)
                    debugCallback(debug);
            });
        }

        easyEws.expandGroup = function (group, successCallback, errorCallback, debugCallback) {
            /// <summary>
            /// PUBLIC:  Expands a group and returns all the members
            /// NOTE:    Does not enumerate groups in groups
            /// </summary>
            /// <param name="group" type="String">The alias for the group to be expanded</param>
            /// <param name="successCallback" type="Function">Callback with an array of aliases - function(String[]) { }</param>
            /// <param name="errorCallback" type="Function">Error handler callback - function(Error) { }</param>
            /// <param name="debugCallback" type="Function">Debug handler returns raw XML - function(String) { }</param>
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
            }, function (debug) {
                if (debugCallback != null)
                    debugCallback(debug);
            });
        };


        easyEws.findConversationItems = function (conversationId, successCallback, errorCallback, debugCallback) {
            /// <summary>
            /// PUBLIC: Find a given conversation by the ID
            /// NOTE: Search for parent:
            ///       http://stackoverflow.com/questions/19008696/exchange-find-items-in-ews-conversation-using-xml-request
            ///       http://www.outlookcode.com/codedetail.aspx?id=1714
            ///       https://msdn.microsoft.com/en-us/library/office/dn610351(v=exchg.150).aspx
            /// </summary>
            /// <param name="conversationId" type="String">The conversation to find</param>
            /// <param name="successCallback" type="Function">Callback with an array of ids - function(String[]) { }</param>
            /// <param name="errorCallback" type="Function">Error handler callback - function(Error) { }</param>
            /// <param name="debugCallback" type="Function">Debug handler returns raw XML - function(String) { }</param>
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
                try {
                    if (xmlDoc == null || xmlDoc.getElementsByTagName("ItemId").length == 0) {
                        if (errorCallback != null)
                            errorCallback(new Error("Invalid XML returned from the server"));
                    } else {
                        $.each(xmlDoc.getElementsByTagName("ItemId"), function (index, value) {
                            returnArray.push(value.getAttribute("Id"));
                        });
                        successCallback(returnArray);
                    }
                } catch (error) {
                    if (errorCallback != null)
                        errorCallback(error);
                }
            }, function (errorDetails) {
                if (errorCallback != null)
                    errorCallback(errorDetails);
            }, function (debug) {
                if (debugCallback != null)
                    debugCallback(debug);
            });
        };

        easyEws.getEwsHeaders = function (itemId, successCallback, errorCallback, debugCallback) {
            /// <summary>
            /// PUBLIC: Gets Internet headers for a spific item
            /// NOTE: https://msdn.microsoft.com/en-us/library/office/aa566013(v=exchg.150).aspx
            /// </summary>
            /// <param name="itemId" type="String">The item ID to get</param>
            /// <param name="successCallback" type="Function">Callback with a Dictionary(key,value) containing the message headers - function(Dictionary) { }</param>
            /// <param name="errorCallback" type="Function">Error handler callback - function(Error) { }</param>
            /// <param name="debugCallback" type="Function">Debug handler returns raw XML - function(String) { }</param>
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
                try {
                    $.each(xmlDoc.getElementsByTagName("InternetMessageHeader"), function (index, value) {
                        returnArray.add(value.getAttribute("HeaderName"), value.textContent);
                    });
                    successCallback(returnArray);
                } catch (error) {
                    if (errorCallback != null)
                        errorCallback(error);
                }
            }, function (errorDetails) {
                if (errorCallback != null)
                    errorCallback(errorDetails);
            }, function (debug) {
                if (debugCallback != null)
                    debugCallback(debug);
            });
        };

        easyEws.updateFolderProperty = function (folderId, propName, propValue, successCallback, errorCallback, debugCallback) {
            /// <summary>
            /// PUBLIC: Updates a folder property. If the property does not exist, it will be created.
            /// </summary>
            /// <param name="folderId" type="String">The ID for the folder</param>
            /// <param name="propName" type="String">The property on the folder to set</param>
            /// <param name="propValue" type="String">The value for the property</param>
            /// <param name="successCallback" type="Function">Callback with the string 'succeeeded' if successful - function(String) { }</param>
            /// <param name="errorCallback" type="Function">Error handler callback - function(Error) { }</param>
            /// <param name="debugCallback" type="Function">Debug handler returns raw XML - function(String) { }</param>
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
            asyncEws(soap, function (data) {
                if (successCallback != null)
                    successCallback('succeeeded');
            }, function (error) {
                if (errorCallback != null)
                    errorCallback(error);
            }, function (debug) {
                if (debugCallback != null)
                    debugCallback(debug);
            });
        }

        // 
        // RETURNS: property value if process completed successfully
        easyEws.getFolderProperty = function (folderId, propName, successCallback, errorCallback, debugCallback) {
            /// <summary>
            /// PUBLIC: Gets a folder property
            /// </summary>
            /// <param name="folderId" type="String">The ID for the folder</param>
            /// <param name="propName" type="String">The property to get</param>
            /// <param name="successCallback" type="Function">Callback with the string value of the property - function(String) { }</param>
            /// <param name="errorCallback" type="Function">Error handler callback - function(Error) { }</param>
            /// <param name="debugCallback" type="Function">Debug handler returns raw XML - function(String) { }</param>
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
            asyncEws(soap, function (xmlDoc) {
                successCallback(xmlDoc.getElementsByTagName("t:Value")[0].textContent);
            }, function (error) {
                if (errorCallback != null)
                    errorCallback(error);
            }, function (debug) {
                if (debugCallback != null)
                    debugCallback(debug);
            });
        }

        easyEws.getFolderId = function (folderName, successCallback, errorCallback, debugCallback) {
            /// <summary>
            /// PUBLIC: Gets the folder id by the given name from the store
            /// </summary>
            /// <param name="folderName" type="String">Name of the folder to get the ID for</param>
            /// <param name="successCallback" type="Function">Callback with the folder ID - function(String) { }</param>
            /// <param name="errorCallback" type="Function">Error handler callback - function(Error) { }</param>
            /// <param name="debugCallback" type="Function">Debug handler returns raw XML - function(String) { }</param>
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
            }, function (debug) {
                if (debugCallback != null)
                    debugCallback(debug);
            });
        }
    }

    return easyEws;

})();

////////////////////////////////////////////////
////////////////////////////////////////////////
easyEws.initialize();    // initialize the class
////////////////////////////////////////////////
////////////////////////////////////////////////


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

