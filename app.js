/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    console.log('aaaa');
    $(document).ready(function () {
      $('#run').click(run);
      sendRequest();
    });
  };

  function run() {
    //Office.context.ui.closeContainer();
    var mailbox = Office.context.mailbox;
    // mailbox.getCallbackTokenAsync(cb);
    // console.log('url:'+mailbox.ewsUrl)
    // console.log('url:'+mailbox.restUrl)
    sendGetRequest();
    
    function cb(asyncResult) {
      var token = asyncResult.value;
      console.log('token: '+token)
    }
    /**
     * Insert your Outlook code here
     */
    console.log('dddd');
  };

  var EventItem = {
    subject: '',
    start: '',
    end: '',
    id: '',
    changeKey: '',
    description: '',
    attendee: [],
  }
  var EWSTool = {

    getXmlTemplate: function(body, timezone){
      var timezoneNode = ''
      if (timezone) {
        timezoneNode = '<t:TimeZoneContext>'+
          '<t:TimeZoneDefinition Id="' + timezone + '" />'+
        '</t:TimeZoneContext>'
      }
      var result = '<?xml version="1.0" encoding="utf-8"?>' +
      '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
            'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
            'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
            'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '<soap:Header>' +
          '<t:RequestServerVersion Version="Exchange2013" />' +
          timezoneNode +
        '</soap:Header>' +
        '<soap:Body>' +
          body +
        '</soap:Body>' +
      '</soap:Envelope>';
  
      return result;
    },

    // Return xml request to get Folder ID
    getFolderXmlRequest: function() {
      var body = '<m:GetFolder>'+
        '<m:FolderShape>'+
          '<t:BaseShape>IdOnly</t:BaseShape>'+
        '</m:FolderShape>'+
        '<m:FolderIds>'+
            '<t:DistinguishedFolderId Id="calendar" />'+
        '</m:FolderIds>'+
      '</m:GetFolder>';
      var result = this.getXmlTemplate(body);
      return result;
    },

    // Return xml query to search appointment with condition
    getListAppointmentXmlRequest: function() {
      var body = '<m:FindItem Traversal="Shallow">'+
        '<m:ItemShape>'+
          '<t:BaseShape>IdOnly</t:BaseShape>'+
          '<t:AdditionalProperties>'+
            '<t:FieldURI FieldURI="item:Subject" />'+
            '<t:FieldURI FieldURI="calendar:Start" />'+
            '<t:FieldURI FieldURI="calendar:End" />'+
          '</t:AdditionalProperties>'+
        '</m:ItemShape>'+
        '<m:CalendarView MaxEntriesReturned="5" StartDate="2018-02-21T17:30:24.127Z" EndDate="2018-03-20T17:30:24.127Z" />'+
        '<m:ParentFolderIds>'+
          '<t:FolderId Id="AQMkADAwATNiZmYAZC1jMG" ChangeKey="AgAAABYAAABBUcKcR" />'+
        '</m:ParentFolderIds>'+
      '</m:FindItem>';
  
      var result = this.getXmlTemplate(body);
      return result;
    },

    // Return xml to create new appointment
    createNewAppointmentXmlRequest: function() {
      var body = '<m:CreateItem SendMeetingInvitations="SendToNone">'+
        '<m:Items>'+
          '<t:CalendarItem>'+
            '<t:Subject>Tennis lesson</t:Subject>'+
            '<t:Body BodyType="HTML">Focus on backhand this week.</t:Body>'+
            '<t:ReminderDueBy>2018-03-09T14:37:10.732-07:00</t:ReminderDueBy>'+
            '<t:Start>2018-03-09T13:00:00.000Z</t:Start>'+
            '<t:End>2018-03-09T15:00:00.000Z</t:End>'+
            '<t:Location>Tennis club</t:Location>'+
            '<t:MeetingTimeZone TimeZoneName="Pacific Standard Time" />'+
          '</t:CalendarItem>'+
        '</m:Items>'+
      '</m:CreateItem>';
      var timezoneName = 'Pacific Standard Time'
      var result = this.getXmlTemplate(body, timezoneName);
      return result;
    },

    updateAppointmentXmlRequest: function() {
      var body = '<UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AutoResolve" SendMeetingInvitationsOrCancellations="SendToNone" '+
                      'xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">'+
            '<ItemChanges>'+
              '<t:ItemChange>'+
                '<t:ItemId Id="AQMkADAwATNiZmYAZdP5UGzCD8TwOec4gAAAMQUBHQAAAA=" ChangeKey="DwAAABYAAAPxPA55ziAADELDp1"/>'+
                '<t:Updates>'+
                  '<t:SetItemField>'+
                    '<t:FieldURI FieldURI="item:Subject" />'+
                      '<t:CalendarItem>'+
                        '<t:Subject>Tennis Lesson moved zzz</t:Subject>'+
                      '</t:CalendarItem>'+
                  '</t:SetItemField>'+
                '</t:Updates>'+
              '</t:ItemChange>'+
            '</ItemChanges>'+
          '</UpdateItem>';
        var result = this.getXmlTemplate(body);
      return result;
    },

    deleteAppointmentXmlRequest: function() {
      var body = '<DeleteItem DeleteType="MoveToDeletedItems" SendMeetingCancellations="SendToAllAndSaveCopy" ' +
      'xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">'+
          '<ItemIds xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
            '<t:ItemId Id="AQMkADAwATNiZmYAZC1jMGE5LT5UGzCD8TwOec4gAAAgENAAAAQVHCnAdP5UGzCD8TwOec4gAAAMQUBHQAAAA=" ChangeKey="DwAAiAADELDp2"/>'+
          '</ItemIds>' +
        '</DeleteItem>';
        var result = this.getXmlTemplate(body);
      return result;
    },
  };

  function deleteAppoiment() {
    // Return a GetItem operation request for the subject of the specified item. 
    var result = '<?xml version="1.0" encoding="utf-8"?>'+
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" '+
           'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'+
      '<soap:Header>'+
        '<t:RequestServerVersion Version="Exchange2013" />'+
        '<t:TimeZoneContext>'+
          '<t:TimeZoneDefinition Id="Pacific Standard Time" />'+
        '</t:TimeZoneContext>'+
      '</soap:Header>'+
      '<soap:Body>'+
        '<m:DeleteItem DeleteType="MoveToDeletedItems" SendMeetingCancellations="SendToAllAndSaveCopy">'+
          '<m:ItemIds>'+
            '<t:ItemId Id="AQMkADAwATNiZmYAZC1jMGE5LTCnAdP5UGzCD8TwOec4gAAAMQUBHQAAAA=" ChangeKey="DwAAABYAAABBUDELDp2" />'+
          '</m:ItemIds>'+
        '</m:DeleteItem>'+
      '</soap:Body>'+
    '</soap:Envelope>';

    return result;
  };
 
 
 
 function sendRequest() {
    // Create a local variable that contains the mailbox.
    var mailbox = Office.context.mailbox;
 
    mailbox.makeEwsRequestAsync(EWSTool.getFolderXmlRequest(), callback);
    //mailbox.makeEwsRequestAsync(EWSTool.createNewAppointmentXmlRequest(), callback);
    mailbox.makeEwsRequestAsync(EWSTool.getListAppointmentXmlRequest(), callback);
 }
 
 function sendGetRequest() {
  // Create a local variable that contains the mailbox.
  var mailbox = Office.context.mailbox;

  mailbox.makeEwsRequestAsync(deleteAppoiment(), callback);
}

 function callback(asyncResult)  {
    var result = asyncResult.value;
    var context = asyncResult.context;
    console.log(result);
    console.log(asyncResult.status);
 
    // Process the returned response here.
 }

})();
