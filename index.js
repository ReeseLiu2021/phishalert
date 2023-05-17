/// <reference path="../App.js" />

(function () {
  'use strict'

  // default confirmation message
  var KB4_DEFAULT_CONFIRMATION_MSG = 'Are you sure you want to report this as a phishing email?'
  var KB4_REPORT_EMAIL_US = 'PHISHALERT@KB4.IO'
  var KB4_REPORT_EMAIL_EU = 'PHISHALERT@KNOWBE4.CO.UK'
  var KB4_RESPONSE_CRID_NOT_FOUND = 'CRID NOT GIVEN'
  var KB4_TASKPANE_URL_MARKER = 'TASKPANEHOME.HTML'

  // The following should have been CONST or LET declarations, but some old IE versions don't allow these keywords.
  var EMAIL_CONTENT_TYPE_MARKER = 'CONTENT-TYPE'
  var EMAIL_WINMAIL_MARKER = 'WINMAIL.DAT'
  var EMAIL_CONTENT_TRANSFER_MARKER = 'CONTENT-TRANSFER-ENCODING'
  var EMAIL_BINARY_CONTENT_MARKER = 'BINARY'
  var EMAIL_MIME_VERSION_MARKER = 'MIME-VERSION'

  var OK_DEFAULT_TEXT = 'Ok'
  var PABV1_COMPLETE_NO_HEADERS = 0
  var PABV1_COMPLETE_NO_FORWARD_ADDRESS = 1

  var PABV1_STATUS_PROCESSING = 0

  var EWS_MAX_SIZE_LIMIT_ASSUMED = 750000 // 3/4 of 1M, assumed as thoretical limit of EWS transactions. (1M is assumed at 1,000,000)
  var isTotalItemSizeBeyondLimit = false

  var KB4_SETTINGS_PULL_INTERVAL_IN_SECONDS = 60 // Will pull settings once every 60 seconds.
  var KB4_SETTINGS_DEFAULT_DISPLAY_DURATION = 3000

  var forwardingData
  var environmentData
  var userEmailAddress
  var currentMessageSender = ''
  var changeKey // retrieved in getItemDataCallback, used for deleting the email
  var headerText
  var _settings
  var show_message_report_pst // bool; show 'congratulations' prompt for submitting
  var message_report_pst // congratulations prompt text
  var show_message_report_non_pst // bool; show 'congratulations' prompt for submitting simulated phishing atttempt
  var message_report_non_pst // congratulations prompt text for simulated phishing atttempt
  var report_button_text
  var ok_button_text = OK_DEFAULT_TEXT
  var report_group_text
  var attachmentTooLargePrefix = 'Email Content Too Large To Send'
  var simulatedPhishing
  var confirmation_message = KB4_DEFAULT_CONFIRMATION_MSG
  var pab_localized_terms

  var KB4_REPORT_EMAIL_US = 'PHISHALERT@KB4.IO'
  var KB4_REPORT_EMAIL_EU = 'PHISHALERT@KNOWBE4.CO.UK'

  // The following should have been CONST or LET declarations, but some old IE versions don't allow these keywords.
  var EMAIL_CONTENT_TYPE_MARKER = 'CONTENT-TYPE'
  var EMAIL_WINMAIL_MARKER = 'WINMAIL.DAT'
  var EMAIL_CONTENT_TRANSFER_MARKER = 'CONTENT-TRANSFER-ENCODING'
  var EMAIL_BINARY_CONTENT_MARKER = 'BINARY'
  var EMAIL_MIME_VERSION_MARKER = 'MIME-VERSION'

  // AsyncDialog variables.
  var notifDialog
  var isToDeleteEmail = false
  var isCloseDialog = false
  var timeout_report_pst = ''
  var displayDurationInMs = KB4_DEFAULT_CONFIRMATION_MSG
  var PABV1_KB4_URLMARKER = 'APPREAD/HOME/'
  var enable_forwarding = false

  function ForwardingData(email, subject) {
    this.email = email
    this.subject = subject
  };

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    var userProfile = Office.context.mailbox.userProfile
    var item = Office.cast.item.toItemRead(Office.context.mailbox.item)
    // Record email details first
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      currentMessageSender = Office.cast.item.toMessageRead(item).from
    } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      currentMessageSender = Office.cast.item.toAppointmentRead(item).organizer
    }
    userEmailAddress = userProfile.emailAddress

    // Initialization of add-in, acquiring KB4-based settings.
    app.initialize()
    getEnvironmentData()
    getAddinSettings()
    $(document).ready(function () {
      displayItemDetails()
    })
  }

  /**
   * Requests configuration data from the KnowBe4 server, and stores it in the Office.roamingsettings facility.
   * 
   * @return {void} 
   */
  function getAddinSettings() {
    var lastSettingsCallDateTime = null
    var tmpLocalterms = null
    var callWebServer = true
    var licensekey = getLicenseKey()
    var serverLocation = getServerLocation()

    // Initialize instance variables to access API objects.
    _settings = Office.context.roamingSettings

    // If the time of the last call to the web server to get user settings is older than 60 seconds, '
    // do a fresh call to the webserver
    lastSettingsCallDateTime = _settings.get('lastSettingsCallDateTime')
    // Force a getaddin settings if the localterms is also null.
    tmpLocalterms = _settings.get('pab_localized_terms')

    if (lastSettingsCallDateTime == null || tmpLocalterms == null) {
      callWebServer = true
    } else {
      // Do date/time comparison
      try {
        // Set a date and get the milliseconds
        var settingsDateTime = new Date(lastSettingsCallDateTime)
        var currentDateTime = new Date(Date())
        // Get the difference in milliseconds.
        var interval = Math.floor((currentDateTime.getTime() - settingsDateTime.getTime()) / 1000)
        callWebServer = (interval > KB4_SETTINGS_PULL_INTERVAL_IN_SECONDS)
      } catch (e) {
        callWebServer = true
      }
    }

    if (callWebServer == true) {
      $.ajax({
        type: 'POST',
        url: serverLocation + '/api/v1/phishalert/addin_initialized',
        async: true,
        context: this,
        headers: { 'Access-Control-Allow-Origin': '*' },
        data: {
          addin_version: encodeURI(environmentData.addin_version),
          auth_token: encodeURI(licensekey),
          os_name: encodeURI(environmentData.os_name),
          os_version: encodeURI(environmentData.os_version),
          os_architecture: encodeURI(environmentData.os_architecture),
          os_locale: encodeURI(environmentData.os_locale),
          outlook_version: environmentData.outlook_version,
          machine_guid: encodeURI(environmentData.machine_guid),
          sender_email: encodeURI(userEmailAddress)
        },
        success: function (data) {
          var myData = JSON.parse(data)

          show_message_report_pst = myData.data.show_message_report_pst
          message_report_pst = myData.data.message_report_pst
          show_message_report_non_pst = myData.data.show_message_report_non_pst
          message_report_non_pst = myData.data.message_report_non_pst
          report_button_text = myData.data.report_button_text
          report_group_text = myData.data.report_group_text
          confirmation_message = myData.data.confirmation_message || KB4_DEFAULT_CONFIRMATION_MSG
          timeout_report_pst = myData.data.timeout_report_pst
          enable_forwarding = myData.data.enable_forwarding

          pab_localized_terms = myData.data.PABLocalizedTerms
          attachmentTooLargePrefix = pab_localized_terms.CommonTerms.TooLargePrefix
          ok_button_text = pab_localized_terms.CommonTerms.Ok
          // Initialize notification mechanism of PAB-Exchange
          app.setupNotificationBlob(pab_localized_terms)
          $('th#subjectLabel').text(pab_localized_terms.CommonTerms.Subject + ':')
          $('th#fromLabel').text(pab_localized_terms.CommonTerms.From + ':')
          $('#buttonReport').prop('value', report_button_text)
          $('#apptitle').text(report_group_text)
          $('#confirmation-message').text(confirmation_message)

          // save settings
          _settings.set('lastSettingsCallDateTime', Date())
          _settings.set('show_message_report_pst', show_message_report_pst)
          _settings.set('message_report_pst', message_report_pst)
          _settings.set('show_message_report_non_pst', show_message_report_non_pst)
          _settings.set('message_report_non_pst', message_report_non_pst)
          _settings.set('report_button_text', report_button_text)
          _settings.set('report_group_text', report_group_text)
          _settings.set('confirmation_message', confirmation_message)
          _settings.set('timeout_report_pst', timeout_report_pst)
          _settings.set('pab_localized_terms', JSON.stringify(pab_localized_terms))
          _settings.set('enable_forwarding', enable_forwarding ? 'true':'')

          // Save roaming settings for the mailbox to the server so that they will be available in the next session.
          _settings.saveAsync(saveMyAddInSettingsCallback)
        },
        error: function (xhr, status, error) {
          endLoaderAndShowError('[getAddinSettings] : ' + xhr.responseText)
          return
        },
        complete: function () {
          // nothing to do here.
        }
      })
    } else {
      show_message_report_pst = _settings.get('show_message_report_pst')
      message_report_pst = _settings.get('message_report_pst')
      show_message_report_non_pst = _settings.get('show_message_report_non_pst')
      message_report_non_pst = _settings.get('message_report_non_pst')
      report_button_text = _settings.get('report_button_text')
      report_group_text = _settings.get('report_group_text')
      confirmation_message = _settings.get('confirmation_message') || KB4_DEFAULT_CONFIRMATION_MSG
      timeout_report_pst = _settings.get('timeout_report_pst')
      enable_forwarding = _settings.get('enable_forwarding') ? true:false

      pab_localized_terms = JSON.parse(_settings.get('pab_localized_terms'))
      attachmentTooLargePrefix = pab_localized_terms.CommonTerms.TooLargePrefix
      ok_button_text = pab_localized_terms.CommonTerms.Ok
      
      // Initialize notification mechanism of PAB-Exchange
      app.setupNotificationBlob(pab_localized_terms)
      $('th#subjectLabel').text(pab_localized_terms.CommonTerms.Subject + ':')
      $('th#fromLabel').text(pab_localized_terms.CommonTerms.From + ':')
      $('#buttonReport').prop('value', report_button_text)
      $('#apptitle').text(report_group_text)
      $('#confirmation-message').text(confirmation_message)
    }
    
    return
  }

  /**
   * Function handler to handle calls to save addin settings into Office-provided roaming settings facility.
   * 
   * @param {object} asyncResult data blob containing the result of the saveSettings() call.
   * 
   * @return {void} 
   */
  function saveMyAddInSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      // Nothing to do here. Allow the next invoke to request the addin settings.
    }
    return
  }

  /**
   * Function handler when displayed name in the "From" entry is clicked.
   * 
   * @return {void} 
   */
  function displayItemDetails() {
    // Displays the "Subject" and "From" fields, based on the current mail item
    var item = Office.cast.item.toItemRead(Office.context.mailbox.item)
    $('#subject').text(item.subject)
    if (currentMessageSender) {
      $('#from').text(currentMessageSender.displayName)
      $('#from').click(function () {
        app.showNotification(currentMessageSender.displayName, currentMessageSender.emailAddress)
      })
    }
    $('#buttonReport').click(submitPhishingReport)
    return
  };


  /**
   * Hide the loader or spinner image.
   * 
   * @return {void} 
   */
  function hideLoader() {
    if (window.location.href.toUpperCase().indexOf(KB4_TASKPANE_URL_MARKER) > -1) {
      $('#pabSpinner').hide();
    } else {
      $('#pabv1SpinnerImage').hide();
    }
    return
  }

  /**
   * Show the loader or spinner image.
   * 
   * @return {void} 
   */
  function showLoader() {
    if (window.location.href.toUpperCase().indexOf(KB4_TASKPANE_URL_MARKER) > -1) {
      $('#pabSpinner').show();
    } else {
      $('#pabv1SpinnerImage').show();
    }
    return
  }

  /**
   * Hide the loader or spinner and then show an error message.
   * 
   * @return {void} 
   */
  function endLoaderAndShowError(errorMessage) {
    hideLoader()
    app.showError(errorMessage)
    return
  }

  /**
   * Hide the loader or spinner and then show a message saying the process was completed.
   * 
   * @return {void} 
   */
  function endLoaderAndShowComplete(completionType, completionMessage) {
    hideLoader()
    app.showComplete(completionType, completionMessage)
    return
  }

  /**
   * Hide the loader or spinner and then show the success message.
   * 
   * @return {void} 
   */
  function endLoaderAndShowSuccess(successMessage) {
    hideLoader()
    app.showSuccess(successMessage)
    return
  }

  /**
   * Function handler when clicking the PAB-report button
   * 
   * @return {void} 
   */
  function submitPhishingReport() {
    app.showStatus(PABV1_STATUS_PROCESSING)
    showLoader()
    // Disable the controls while sending data
    $('#buttonReport').prop('disabled', true)
    sendHeadersRequest()
    return
  };

  /**
   * Creates the SOAP request envelope, wrapping the raw EWS request.
   * 
   * @param {string} request string representing the raw EWS request.
   * 
   * @return {string} XML/SOAP envelope used for all EWS requests
   */
  function getSoapEnvelope(request) {
    // Wrap an Exchange Web Services request in a SOAP envelope.
    var result =
      "<?xml version='1.0' encoding='utf-8'?>" +
      "<soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'" +
      "               xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types'>" +
      '  <soap:Header>' +
      "     <t:RequestServerVersion Version='Exchange2013'/>" +
      '  </soap:Header>' +
      '  <soap:Body>' +
      request +
      '  </soap:Body>' +
      '</soap:Envelope>'
    return result
  };

  /**
  * Creates the SOAP request to get the headers of an email
  * 
  * @param {string} id email id whose headers will be requested 
  * 
  * @return {string} XML/SOAP request for use in asking for the email headers.
  */
  function getHeadersRequest(id) {
    // Return a GetItem EWS operation request for the headers of the specified item.
    var result =
      "    <GetItem xmlns='http://schemas.microsoft.com/exchange/services/2006/messages'>" +
      '      <ItemShape>' +
      '        <t:BaseShape>IdOnly</t:BaseShape>' +
      '        <t:BodyType>Text</t:BodyType>' +
      '        <t:AdditionalProperties>' +
      '         <t:FieldURI FieldURI="item:Size" />' +
      // PR_TRANSPORT_MESSAGE_HEADERS
      "         <t:ExtendedFieldURI PropertyTag='0x007D' PropertyType='String' />" +
      '        </t:AdditionalProperties>' +
      '      </ItemShape>' +
      "      <ItemIds><t:ItemId Id='" + id + "'/></ItemIds>" +
      '    </GetItem>'
    return result
  };


  /**
  * Executes the actual http-request to ask for the headers of the email about to be reported.
  * 
  * @return {void} 
  */
  function sendHeadersRequest() {
    var mailbox = Office.context.mailbox
    var request = getHeadersRequest(mailbox.item.itemId)
    var envelope = getSoapEnvelope(request)
    try {
      mailbox.makeEwsRequestAsync(envelope, getHeadersCallback)
    } catch (e) {
      endLoaderAndShowError('[sendHeadersRequest] : ' + e)
    }
    return
  }

  // This function plug in filters nodes for the one that matches the given name.
  // This sidesteps the issues in jquery"s selector logic.
  (function ($) {
    $.fn.filterNode = function (node) {
      return this.find('*').filter(function () {
        return this.nodeName === node
      })
    }
  })(jQuery)

  /**
  * Function called when the EWS request to get headers is complete.
  * 
  * @param {object} asyncResult blob of data representing the result of the getHeaders() call.
  * 
  * @return {void} 
  */
  function getHeadersCallback(asyncResult) {
    // Process the returned response here.
    if (asyncResult.value) {
      var prop = null
      var secondProp = null
      try {
        var response = $.parseXML(asyncResult.value)
        var responseDOM = $(response)
        if (responseDOM) {
          // See http://stackoverflow.com/questions/853740/jquery-xml-parsing-with-namespaces
          // See also http://www.steveworkman.com/html5-2/javascript/2011/improving-javascript-xml-node-finding-performance-by-2000
          // We can do this because we know there's only the one property.
          var itemSize = responseDOM.filterNode('t:Size')[0]
          var actualSize = 0
          try {
            actualSize = parseInt(itemSize.textContent)
            if (actualSize > EWS_MAX_SIZE_LIMIT_ASSUMED) {
              isTotalItemSizeBeyondLimit = true
            }
          } catch (errVal) {
            // nothing to do here.
          }
          prop = responseDOM.filterNode('t:ExtendedProperty')[0]
          // let us get the changeKey here at this point.
          secondProp = responseDOM.filterNode('t:ItemId')[0]
          changeKey = secondProp.getAttribute('ChangeKey')
        }
      } catch (e) {
        // Nothing to do here.
      }

      if (!prop) {
        // Unlike the old implementation which showed granular messaging, this time we only show that the email they are forwarding has no headers.
        endLoaderAndShowComplete(PABV1_COMPLETE_NO_HEADERS)
        setTimeout(function () { deleteEmail(false) }, KB4_SETTINGS_DEFAULT_DISPLAY_DURATION)
        return
      }

      headerText = prop.textContent.toString()
      simulatedPhishing = (headerText.toLowerCase().indexOf('x-phish-crid') > -1)

      // NOTE Step 2: Send headers or get forwarding address
      if (simulatedPhishing) {
        // submit simulated phishing report: send crid value to web service
        var headerArray
        var crid
        try {
          // BUG This regEx doesn't work in JavaScript
          // regEx = '^(?<header_key>[-A-Za-z0-9]+)(?<seperator>:[ \t]*)' + '(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)';
          // headerArray = headerText.match(regEx); //https://msdn.microsoft.com/en-us/library/7df7sf4x(v=vs.94).aspx
          // Split headers into array
          headerArray = headerText.split('\n') // This is ugly but can get 'er done - some header values may split across array elements, but we just need to look for one header
          for (var index = 0; index < headerArray.length; index++) {
            if (headerArray[index].toLowerCase().indexOf('x-phish-crid') > -1) {
              var headerlinevals = headerArray[index].split(':')
              crid = headerlinevals[1]
              crid = crid.replace(/\r?\n|\r|\s+/g, '')
            }
          }
        } catch (e) {
          endLoaderAndShowError('[getHeadersCallback] : ' + e)
          return
        }

        // NOTE Send header to web service
        submitHeader(crid)
      } else {
        // submit non-simulated phishing report: forward email with original headers
        getForwardingEmailAddress(headerText)
      }
    } else if (asyncResult.error) {
      endLoaderAndShowError('[getHeadersCallback] : ' + asyncResult.error.message)
      return
    }
    return
  };


  /**
   * Submit/Report the campaign id to the KMSAT server and display the appropriate success message.
   * 
   * @param {string} crid the campaign id obtained from the email headers.
   * 
   * @return {void} 
   */
  function submitHeader(crid) {
    var licensekey = getLicenseKey()
    var serverLocation = getServerLocation()
    var errText = ''
    try {
      $.ajax({
        type: 'POST',
        url: serverLocation + '/api/v1/phishalert/report',
        async: true,
        context: this,
        headers: { 'Access-Control-Allow-Origin': '*' },
        data: { crid: crid, addin_version: encodeURI(environmentData.addin_version), auth_token: encodeURI(licensekey), os_name: encodeURI(environmentData.os_name), os_version: encodeURI(environmentData.os_version), os_architecture: encodeURI(environmentData.os_architecture), os_locale: encodeURI(environmentData.os_locale), outlook_version: environmentData.outlook_version, machine_guid: encodeURI(environmentData.machine_guid), sender_email: encodeURI(userEmailAddress) },
        success: function (data) {
          // If the success message is to be displayed then we handle it here.
          if (show_message_report_pst) {
            try {
              hideLoader()
              try {
                displayDurationInMs = parseInt(timeout_report_pst) * 1000
              } catch (e) {
                displayDurationInMs = KB4_SETTINGS_DEFAULT_DISPLAY_DURATION // Agreed delay whether error or not is 3 seconds.
              }
              openDialogAsIframe(ok_button_text, message_report_pst, displayDurationInMs)
            } catch (e) {
              endLoaderAndShowSuccess(message_report_pst);
              setDismissal()
            }
          } else {
            // Get A changekey item and delete email.
            getChangeKeyAndDeleteEmail(false)
          }
        },
        error: function (xhr, status, error) {
          // If the axios client fails (whether status is 400 or not), we forward the email.
          getForwardingEmailAddress(headerText)
        }
      })
    } catch (postReportEx) {
      // Get A changekey item and delete email.
      getChangeKeyAndDeleteEmail(false)
    }
  }

  /**
   * Get the user license based on the url
   * 
   * @return {string} the user license.`
   */
  function getLicenseKey() {
    var result
    // Get license key from manifest URL
    var res = window.location.href.match(/\/phishalertonline\/(.*?)(\/.*?)/)
    if (res == null) { return null }
    if (res.length >= 2) {
      result = res[1]
    }
    return result
  }

  /**
   * Get the server location based on the url on the html window.
   * 
   * @return {string} the server url
   */
  function getServerLocation() {
    var result
    var res = window.location.href.match(/(.*?)\/phishalertonline\/(.*?)(\/.*?)/)
    if (res == null) { return null }
    if (res.length >= 2) {
      result = res[1]
    }
    return result
  }

  /**
   * Environment data initialization.
   * 
   */
  function EnvironmentData(addin_version, os_name, os_version, os_architecture, os_locale, outlook_version, machine_guid) {
    this.addin_version = addin_version
    this.os_name = os_name
    this.os_version = os_version
    this.os_architecture = os_architecture
    this.os_locale = os_locale
    this.outlook_version = outlook_version
    this.machine_guid = machine_guid
  };


  /**
   * Sets the environment data like locale, OS, version and others.
   * 
   * @return {void} 
   */
  function getEnvironmentData() {
    var diags = Office.context.mailbox.diagnostics
    var contxt = Office.context
    var oslocale, outlookversion
    try {
      // Get environment settings
      // useragent = $.HTTP_USER_AGENT; //How is this used?
      outlookversion = diags.hostVersion
      oslocale = contxt.displayLanguage
      // window.location.href =  https://localhost:44301/AppRead/Home/Home.html?_host_Info=Outlook|Win32|16.00|en-US
      var paramvalues = window.location.search.split('=')
      var os_architecture = 'n/a'
      if (paramvalues.length > 1) // outlook web sends these params in, it is unreliable though what params are actually sent!
      {
        // hostInfo = paramvalues[1].split('|');
        os_architecture = 'web'// hostInfo[1]
      }
      environmentData = new EnvironmentData()
      environmentData.addin_version = '1.0' // hardcode this for now
      environmentData.os_architecture = os_architecture
      environmentData.outlook_version = outlookversion
      environmentData.os_locale = oslocale
      environmentData.os_name = 'unknown'
      environmentData.os_version = 'unknown'
      environmentData.machine_guid = 'unknown' // E.g: 15096333-e04f-4726-badb-ef151ecd6990
    } catch (e) {
      // Nothing to do here.
    }
    return
  }

  /**
   * Request the email addresses to whom the report email will be sent.
   * 
   * @param {string} headerText list of headers in  the email to be reported.
   * 
   * @return {void} 
   */
  function getForwardingEmailAddress(headerText) {
    // get forwarding data
    var licensekey = getLicenseKey()
    var serverLocation = getServerLocation()
    $.ajax({
      type: 'POST',
      url: serverLocation + '/api/v1/phishalert/forward_emails',
      async: true,
      context: this,
      headers: { 'Access-Control-Allow-Origin': '*' },
      data: { addin_version: encodeURI(environmentData.addin_version), auth_token: encodeURI(licensekey), os_name: encodeURI(environmentData.os_name), os_version: encodeURI(environmentData.os_version), os_architecture: encodeURI(environmentData.os_architecture), os_locale: encodeURI(environmentData.os_locale), outlook_version: environmentData.outlook_version, machine_guid: encodeURI(environmentData.machine_guid), sender_email: encodeURI(userEmailAddress) },
      success: function (data) {
        try {
          var myData = JSON.parse(data)
          var item = Office.cast.item.toItemRead(Office.context.mailbox.item)
          forwardingData = new ForwardingData(myData.data.email_forward, myData.data.email_forward_subject)
          var mailbox = Office.context.mailbox
          mailbox.makeEwsRequestAsync(getItemDataRequest(item.itemId, ['Body', 'Subject', 'MimeContent']), getItemDataCallback, { itemId: item.itemId, headers: headerText })
        } catch (e) {
          endLoaderAndShowError('[getForwardingEmailAddress] : ' + '(' + status + ')' + e)
        }
      },
      error: function (xhr, status, error) {
        endLoaderAndShowError('[getForwardingEmailAddress] : ' + '(' + status + ')' + error)
      },
      complete: function () { }
    })
    return
  };

  /**
   * Creates the SOAP request to get several fields of information on the message ID provided
   * 
   * @param {string} item_id id of the message 
   * @param {string} fields an array of strings representing data we want to gather from the server based on the message-id
   * 
   * @return {string} XML/SOAP request effectively executing a getItem() call on the Exchange Server.
   */
  function getItemDataRequest(item_id, fields) {
    var allFields = ''
    fields.forEach(function (field) {
      allFields += ('<t:FieldURI FieldURI="item:' + field + '" />')
    })
    var request = '<?xml version="1.0" encoding="utf-8"?>' +
      '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
      '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
      '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
      '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
      '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
      '  <soap:Header>' +
      '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
      '  </soap:Header>' +
      '  <soap:Body>' +
      '    <GetItem' +
      '                xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
      '                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
      '      <ItemShape>' +
      '        <t:BaseShape>IdOnly</t:BaseShape>' +
      '           <t:AdditionalProperties>' +
                    allFields +
      '           </t:AdditionalProperties>' +
      '      </ItemShape>' +
      '      <ItemIds>' +
      '        <t:ItemId Id="' + item_id + '"/>' +
      '      </ItemIds>' +
      '    </GetItem>' +
      '  </soap:Body>' +
      '</soap:Envelope>'
    return request
  }


  /**
   * Callback function whenever a request to getItem() for deletion is executed..
   * 
   * @param {object} asyncResult blob of data representing the result of the getItem() call.
   * 
   * @return {void} 
   */
  function getItemToDeleteCallback(asyncResult) {
    if (asyncResult == null) {
      endLoaderAndShowError("[getItemToDeleteCallback] : asyncResult=''")
      return
    }

    if (asyncResult.error != null) {
      endLoaderAndShowError('[getItemToDeleteCallback] : ' + asyncResult.error.message)
      return
    } else {
      var errorMsg
      var prop = null
      try {
        var response = $.parseXML(asyncResult.value)
        var responseDOM = $(response)
        if (responseDOM) {
          prop = responseDOM.filterNode('t:ItemId')[0]
          changeKey = prop.getAttribute('ChangeKey')
        }
      } catch (e) {
        errorMsg = e
      }
      if (!prop) {
        if (errorMsg) {
          endLoaderAndShowError('[getItemToDeleteCallback] : ' + errorMsg)
        } else {
          endLoaderAndShowError("[getItemToDeleteCallback] : prop=''")
        }
      } else {
        if (changeKey) {
          // NOW delete the email
          var mailbox = Office.context.mailbox
          var item = Office.cast.item.toItemRead(Office.context.mailbox.item)
          mailbox.makeEwsRequestAsync(moveItemRequest(item.itemId, changeKey), moveItemCallback)
        } else {
          endLoaderAndShowError('[getItemToDeleteCallback] : ' + prop.textContent)
        }
      }
    }
    return
  }

  /**
   * Create an XML entry representing the email addresses to whom the emails will be sent.
   * 
   * @param {array} emlAddressArray array of string containing the recipient email addresses.
   * 
   * @return {string} an XML string representing the TO/CC/BCC email addresses, for use in the SOAP
   */
  function buildEmailAddressList(emlAddressArray) {
    var finalAddressList = ''
    for (var i = 0; i < emlAddressArray.length; i++) {
      finalAddressList += '<t:Mailbox><t:EmailAddress>' + emlAddressArray[i] + '</t:EmailAddress></t:Mailbox>'
    }
    return finalAddressList
  }

  /**
   * Creates the list of email addresses to send the report to, based on a blob data coming from KB4Server
   * 
   * @param {string} forwardingData a comma-separated string containing a set of email addresses.
   * 
   * @return {object} blob of data containing email addresses to whom the report will be sent.
   */
  function buildForwardingSoapData(forwardingData) {
    var addresses = forwardingData.email.split(',')
    var addressTo = ''
    var addressBcc = ''
    var trimmed_address = ''
    var addressToRecipientArray = []
    var addressBCCRecipientArray = []
    for (var i = 0; i < addresses.length; i++) {
      if (addresses[i] && addresses[i] != '') {
        trimmed_address = addresses[i].trim()
        var upperCaseAddress = trimmed_address.toUpperCase()
        // There could never be a case where both KB4_REPORT_EMAIL_US and KB4_REPORT_EMAIL_EU are both existing.
        // Unless they are manually registered as recipients. However, let us just try adding them both.
        if (upperCaseAddress == KB4_REPORT_EMAIL_US || upperCaseAddress == KB4_REPORT_EMAIL_EU) {
          addressBCCRecipientArray.push(trimmed_address)
        } else {
          addressToRecipientArray.push(trimmed_address)
        }
      }
    }
    // is there a valid "To" Recipient? if so, process normally.
    if (addressToRecipientArray.length > 0) {
      addressTo = buildEmailAddressList(addressToRecipientArray)
      addressBcc = buildEmailAddressList(addressBCCRecipientArray)
    } else {
      // now if there is no "To" recipient, push all bcc recipients (1 per region) to the to "to" recipients
      addressTo = buildEmailAddressList(addressBCCRecipientArray)
    }
    if (addressTo != '') {
      if (addressBcc != '') {
        return {
          'ToRecipients': addressTo,
          'BccRecipients': addressBcc
        };
      }
      return { 'ToRecipients': addressTo };
    }
    return null
  }


  /**
   * Callback function executed when a GettItemData() call is finished.
   * 
   * @param {object} asyncResult a structured data containing information on the result of the getItemData asyncrhonous call.
   * 
   * @return {void} 
   */
  function getItemDataCallback(asyncResult) {
    // We first get the Office context data.
    var mailbox = Office.context.mailbox
    var item = Office.cast.item.toItemRead(Office.context.mailbox.item)
    var addressesSoap = buildForwardingSoapData(forwardingData)
    var isOver1mb = false
    // Something wrong with the email addresses to forward the emails to.
    if (addressesSoap == null) {
      endLoaderAndShowComplete(PABV1_COMPLETE_NO_FORWARD_ADDRESS)
      // Show the "complete" message but display the error for 3 seconds.
      setTimeout(function () { deleteEmail(false) }, KB4_SETTINGS_DEFAULT_DISPLAY_DURATION)
      return
    }

    // Let us mark this as a repeat call.
    if (asyncResult.asyncContext.hasOwnProperty('isOver1mb')) {
      isOver1mb = true
    }

    // We don't have any data from the query, let us just return with a null result.
    if (asyncResult == null) {
      endLoaderAndShowError("[getItemDataCallback] : asyncResult=''")
      return
    }

    // Is the result an error? Is it because we have over 1mb of data?
    if (asyncResult.error != null) {
      // Check to see if request failed because response is too large
      var isErrorMessageOver1MB = (asyncResult.error.message.toUpperCase().indexOf('EXCEEDS 1 MB') !== -1)
      var isError9020 = (asyncResult.error.code == 9020)
      if (!isOver1mb && (isError9020 || isErrorMessageOver1MB)) {
        var headers = asyncResult.asyncContext.headers
        Office.context.mailbox.makeEwsRequestAsync(getItemDataRequest(asyncResult.asyncContext.itemId, ['Subject']), getItemDataCallback, { headers: headers, isOver1mb: 'true' })
        return
      } else {
        // Display an error and then return.
        endLoaderAndShowError('[getItemDataCallback] : ' + asyncResult.error.message)
        return
      }
    }

    // PAB-933-It seems that for MAC-Outlook, EWS requests return success even if the response is beyond 1MB.
    // This is a workaround, to ensure that PAB-Exchange does not hang or get stuck in a loop when this case happens.
    // Until microsoft comes with a fix on their bug.
    if (!isOver1mb && asyncResult.value == null && isTotalItemSizeBeyondLimit) {
      var headers = asyncResult.asyncContext.headers
      Office.context.mailbox.makeEwsRequestAsync(getItemDataRequest(asyncResult.asyncContext.itemId, ['Subject']), getItemDataCallback, { headers: headers, isOver1mb: 'true' })
      return
    }

    // Now let us process the result, at this point this should either be successful or over 1MB
    var errorMsg = null
    var prop = null
    var mimeContent = ''

    try {
      var response = $.parseXML(asyncResult.value)
      var responseDOM = $(response)
      if (responseDOM) {
        prop = responseDOM.filterNode('t:ItemId')[0]
      }
    } catch (e) {
      // Something wrong with the parsing.
      errorMsg = e
    }
    if (!prop) {
      if (errorMsg) {
        endLoaderAndShowError('[getItemDataCallback] : ' + errorMsg)
      } else {
        endLoaderAndShowError("[getItemDataCallback] : (prop='') asyncResult=" + asyncResult.value)
      }
      return
    }
    // Update the value of the changekey, if not, use the one in the sendHeadersRequest()
    changeKey = prop.getAttribute('ChangeKey') // Used to delete the email after it is forwarded
    // Get the MimeContent so we can construct an attachment from the source email and attach it 
    // to the new email we are sending (no longer forwarding the email - just creating a blank one with some properties)
    prop = null
    prop = responseDOM.filterNode('t:MimeContent')[0]
    // let us get the mimecontent if it is avaiblable.
    if (prop != null) {
      mimeContent = prop.textContent
    }

    // NOTE Step 4: Create new email and send with copy of source email as attachment
    var bodyContent = ''
    var isHTML = false // default is plain text only.
    var sourceSubject = !item.subject ? '' : item.subject;

    // mimeContent can be empty because if we failed to retrieve it on the first request it is because
    // the response was larger than 1MB.
    if (isOver1mb) {
      // We will only submit the headers.
      sourceSubject = '[' + attachmentTooLargePrefix + ']' + sourceSubject
      if (isHTML) {
        bodyContent = asyncResult.asyncContext.headers.replace(/(?:\r\n|\r|\n)/g, '<br>')
      } else {
        bodyContent = asyncResult.asyncContext.headers
      }
    } else {
      // Get the body of the email.
      var body = responseDOM.filterNode('t:Body')[0]
      var bodyType = body.getAttribute('BodyType')
      bodyContent = body.textContent
      if (bodyType == 'HTML') {
        isHTML = true
      }
    }
    var xml = ''
    try {
      xml = createAndSendItemWithAttachmentsRequest(forwardingData.subject, sourceSubject, addressesSoap, mimeContent, bodyContent, isHTML)
    } catch (xmlException) {
      endLoaderAndShowError('[getItemDataCallback] : ' + errorMsg)
      return
    }
    try {
      mailbox.makeEwsRequestAsync(xml, createAndSendItemWithAttachmentsCallback, { headers: { 'Content-Type': 'text/xml; charset=utf-8' }}) 
    } catch (xmlException) {
      // Just before we send out this data, let us check the length of the XML 
      if (!isOver1mb && xml.length >= 1000000) {
          var headers = asyncResult.asyncContext.headers
          Office.context.mailbox.makeEwsRequestAsync(getItemDataRequest(asyncResult.asyncContext.itemId, ['Subject']), getItemDataCallback, { headers: headers, isOver1mb: 'true' })
          return
      } else {
        endLoaderAndShowError('[getItemDataCallback] : ' + errorMsg)
        return       
      }
    }
    return
  }

  /**
   * Move an email into the delteditems folder without having to update the changekey.
   * 
   * @param {boolean} isFromAsyncDialog set to true if the call is from after displaying an asynchronous dialog. False if not. 
   * 
   * @return {void} 
   */
  function deleteEmail(isFromAsyncDialog) {
    var mailbox = Office.context.mailbox
    var item = Office.cast.item.toItemRead(Office.context.mailbox.item)
    mailbox.makeEwsRequestAsync(moveItemRequest(item.itemId, changeKey), moveItemCallback, { isFromAsyncDialog: isFromAsyncDialog })
    return
  }


  /**
   * Convenience function to hide notification and delete email after a pre-determined duration.
   * 
   * @return {void} 
   */
  function setDismissal() {
    var dismissed = false
    var dismissBtn = "<span class='dismiss-btn'>" + ok_button_text + "</span>"
    $('#notification-message-body').append(dismissBtn)

    // delete on dismiss click
    $('.dismiss-btn').click(function () {
      dismissed = true
      $('.dismiss-btn').off('click')
      $('.dismiss-btn').addClass('btn-disabled')
      $('#notification-message').hide()
      deleteEmail(false)
    })

    // otherwise delete after 3 seconds
    setTimeout(function () {
      if (dismissed !== true) {
        $('#notification-message').hide()
        deleteEmail(false)
      }
    }, KB4_SETTINGS_DEFAULT_DISPLAY_DURATION)
    return
  }

  /**
   * Callback function executed right after a sendItemWithAttachments is requested via Exchange API
   *
   * @param {object} asyncResult a structured data containing information on the result of the EWS-API request
   * 
   * @return {void} 
   */
  function createAndSendItemWithAttachmentsCallback(asyncResult) {
    if (asyncResult == null) {
      endLoaderAndShowError("[createAndSendItemWithAttachmentsCallback] : asyncResult=''")
      return
    }
    if (asyncResult.error != null) {
      endLoaderAndShowError('[createAndSendItemWithAttachmentsCallback] : ' + asyncResult.error.message)
    } else {
      var errorMsg
      var prop = null
      try {
        var response = $.parseXML(asyncResult.value)
        var responseDOM = $(response)
        if (responseDOM) {
          prop = responseDOM.filterNode('m:ResponseCode')[0]
        }
      } catch (e) {
        errorMsg = e
      }
      if (!prop) {
        if (errorMsg) {
          endLoaderAndShowError('[createAndSendItemWithAttachmentsCallback] : ' + errorMsg)
        } else {
          endLoaderAndShowError("[createAndSendItemWithAttachmentsCallback] :  prop=''")
        }
      } else {
        // NOTE Step 6: Verify forward result
        if (prop.textContent == 'NoError') {
          if (show_message_report_non_pst) {
            try {
              hideLoader()
              try {
                displayDurationInMs = parseInt(timeout_report_pst) * 1000
              } catch (e) {
                displayDurationInMs = KB4_SETTINGS_DEFAULT_DISPLAY_DURATION // Agreed delay whether error or not is 3 seconds.
              }
              openDialogAsIframe(ok_button_text, message_report_non_pst, displayDurationInMs)
            } catch (e) {
              endLoaderAndShowSuccess(message_report_non_pst);
              setDismissal()
            }
          } else {
            deleteEmail(false)
          }
        } else {
          endLoaderAndShowError('[createAndSendItemWithAttachmentsCallback] : ' + prop.textContent)
        }
      }
    }
    return
  }

  /**
   * Create an XML/SOAP request based on information provided. This SOAP string is then sent as part of an HTTPS request.
   *
   * @param {string} usrSubjectPrefix the subject prefix to be used
   * @param {string} usrSourceSubject the original subject of the email to be reported
   * @param {object} addressesSoap a structured data containing the TO, CC, BCC email addresses
   * @param {string} mimeContent Base64 encoded string that is the payload of the email
   * @param {string} usrBody the original body of the email to be reported.
   * @param {boolean} isHTML true if the body of the email is in html format.
   * 
   * @return {string} XML/SOAP packet conformant to the "send-email" request using Exchange Server API.
   */
  function createAndSendItemWithAttachmentsRequest(usrSubjectPrefix, usrSourceSubject, addressesSoap, mimeContent, usrBody, isHTML) {
    // we need to encode subject and body even for the text version, so it is valid xml when calling the EWS       
    var subjectPrefix = escapeStringForXML(usrSubjectPrefix)
    var sourceSubject = escapeStringForXML(usrSourceSubject)
    var body = escapeStringForXML(usrBody)
    // Let us get the message ID.
    var item = Office.cast.item.toItemRead(Office.context.mailbox.item)
    var escapedInternetMessageId = escapeStringForXML(item.internetMessageId)


    // https://msdn.microsoft.com/en-us/library/office/dn726694%28v=exchg.150%29.aspx#bk_createattachews
    // NOTE: We must set an extended property (0x0E07 to 0) on the attachment so it is not in compose mode when opened https://social.msdn.microsoft.com/Forums/en-US/5386612f-d897-458c-a295-c64bee4263bf/message-has-not-been-sent-when-creating-an-email-item?forum=exchangesvrdevelopment
    var bodyContent = ''
    var request = ''

    // We needed to change the code to use old-style string manipulation because IE-JS interpreter does not support new method.
    var headersForForwardIeReferencesAndinReplyTo = ''
    if (enable_forwarding) {
      headersForForwardIeReferencesAndinReplyTo = '<t:ExtendedProperty>\n' +
      '    <t:ExtendedFieldURI PropertyTag="4162" PropertyType="String" />\n' +
      '    <t:Value>' + escapedInternetMessageId + '</t:Value>\n' + 
      '    </t:ExtendedProperty>\n' +
      '    <t:ExtendedProperty>\n' + 
      '    <t:ExtendedFieldURI PropertyTag="4153" PropertyType="String" />\n' +
      '    <t:Value>' + escapedInternetMessageId + '</t:Value>\n' +
      '    </t:ExtendedProperty>'
    }
  
    if (isHTML == true) {
      bodyContent = '<t:Body BodyType="HTML">' + body + '</t:Body>'
    } else {
      bodyContent = '<t:Body BodyType="Text">' + body + '</t:Body>'
    }
    //PAB-598 - It seems that old exchange servers are very strict in the XSD validation, let us correct this.
    var forwardingAddressesString = '	          <t:ToRecipients>	' + addressesSoap.ToRecipients +
      '	          </t:ToRecipients>	'
    if (addressesSoap.BccRecipients) {
      forwardingAddressesString += '	          <t:BccRecipients>	' + addressesSoap.BccRecipients +
        '	          </t:BccRecipients>	'
    }
    request = '<?xml version="1.0" encoding="utf-8"?>' +
      '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
      '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
      '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
      '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
      '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
      '  <soap:Header>' +
      '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
      '  </soap:Header>' +
      '  <soap:Body>' +
      '    <m:CreateItem MessageDisposition="SendAndSaveCopy">' +
      '	      <m:Items>	' +
      '	        <t:Message>	' +
      '	          <t:Subject>' + subjectPrefix + sourceSubject + '</t:Subject>	' +
                  bodyContent +
                  getAttachmentXML(mimeContent) +
                  headersForForwardIeReferencesAndinReplyTo +
                  forwardingAddressesString +
      '	        </t:Message>	' +
      '	      </m:Items>	' +
      '	    </m:CreateItem>	' +
      '  </soap:Body>' +
      '</soap:Envelope>'
    return request
  }


  /**
   * Call back function for the moveItem call.
   *
   * @param {object} asyncResult blob of data containing info on request
   * 
   * @return {void} 
   */
  function moveItemCallback(asyncResult) {
    // Whether moving items to deleted items fail or not, no error will be displayed.
    // This is to ensure that any prior error messages will not be replaced, allowing debug to take place.
    hideLoader()
    try {
      if (asyncResult.asyncContext.isFromAsyncDialog) {
        // Nothing to do here for now.
      }
      if (Office.context.requirements.isSetSupported('Mailbox', '1.5')) {
        Office.context.ui.closeContainer()
      } else {
        $('#notification-message').hide()
      }
    } catch (moveItemException) {
      // nothing to do here. 
    }
    return
  }

  /**
   * Encode a string for XML use.
   *
   * @param {string} html raw string to encode
   *
   * @return {string} XML-escaped version of the original string.
   */
  function escapeStringForXML (html) {
    if (html) {
      // https://www.ibm.com/docs/en/was-liberty/base?topic=manually-xml-escape-characters
      var entityMap = { '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&apos;'
      }
      return String(html).replace(/[&<>"']/g, function (s) {
        return entityMap[s]
      })
    }
    return ''
  }

  /**
   * Form the moveItem request based on the id and key provided.
   *
   * @param {string} item_id id of the message 
   * @param {string} changeKey access granting key.
   * 
   * @return {string} XML/SOAP request to move an email to deleted items.
   */
  function moveItemRequest(item_id, changeKey) {
    var request
    request = '<?xml version="1.0" encoding="utf-8"?> ' +
      '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
      '        xmlns:xsd="http://www.w3.org/2001/XMLSchema" ' +
      '        xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
      '        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
      ' <soap:Header>' +
      '   <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
      ' </soap:Header>' +
      ' <soap:Body> ' +
      '   <MoveItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
      '     xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
      '     <ToFolderId> ' +
      '       <t:DistinguishedFolderId Id="deleteditems"/> ' +
      '     </ToFolderId> ' +
      '     <ItemIds> ' +
      '       <t:ItemId Id="' + item_id + '" ChangeKey="' + changeKey + '"/> ' +
      '     </ItemIds> ' +
      '   </MoveItem> ' +
      ' </soap:Body> ' +
      '</soap:Envelope>'
    return request
  }

  /**
   * HTML Encode a string.
   *
   * @param {string} html raw string to encode
   * 
   * @return {string} htmlEncoded() version of the original string.
   */
  function htmlEncode(html) {
    // create a in-memory div, set it's inner text(which jQuery automatically encodes)
    // then grab the encoded contents back out.  The div never exists on the page.
    var parsed = $('<div/>').text(html).html()
    // added because of an error we are getting from EWS if the body contains an &nbsp;
    parsed = parsed.replace(/&amp;nbsp;|&nbsp;/gi, '&#160;')
    parsed = parsed.replace(/&shy;/gi, '&#173;')
    return parsed
  }

  /**
   * Create a SOAP-XML phrase/blob that creates an eml file attachment in the email. 
   * EML file contains the contents of the email file being reported.
   *
   * @param {obejct} mimeContent the EWS-server provide mime data
   * 
   * @return {string} an xml blob that represents an eml attachment.
   */
  function getAttachmentXML(mimeContent) {
    var xml = ''
    if (mimeContent !== '') {
      // Let us create an eml file first.
      // PAB-760 - Recreate the eml file by using the fully headers and the provided mimecontent.            
      var rebuiltEMLFile = btoa(buildEMLFileFromMimeContentAndFullHeaders(mimeContent, headerText))
      xml = '     <t:Attachments>	' +
        '	            <t:FileAttachment>	' +
        '	              <t:Name>phish_alert_sp1_1.0.0.0.eml</t:Name>	' +
        '	              <t:ContentType>application/octet-stream</t:ContentType>	' +
        '	              <t:IsInline>false</t:IsInline>	' +
        '	              <t:IsContactPhoto>false</t:IsContactPhoto>	' +
        '	              <t:Content>' + rebuiltEMLFile + '</t:Content>	' +
        '	            </t:FileAttachment>	' +
        '	          </t:Attachments>	'
    }
    return xml
  }

  /**
   * Removes the unnecessary content-type and content-transfer encoding from the headerset.
   * EWS Servers, when asked for headers returns "Content-Type: application/ms-tnef; name="winmail.dat" 
   * and "Content-Transfer-Encoding: binary" whenever they get the headers of an email sent via Outlook/OWA 
   *
   * @param {object} usrObject a user object that contains the headers array. This is directly modified within the function.
   * 
   * @return {boolean} true if object was modified or false if not.
   */
  function removeWinMailDatHeader(usrObject) {
    // We need to check for the winmail.dat Content-type
    var retVal = false
    var localHeaderSet = usrObject.headerArray
    for (var idx = 0; idx < localHeaderSet.length; idx++) {
      if (localHeaderSet[idx].toUpperCase().startsWith(EMAIL_CONTENT_TYPE_MARKER)) {
        if (localHeaderSet[idx].toUpperCase().indexOf(EMAIL_WINMAIL_MARKER) > 0) {
          // Replace the header with an empty string.
          localHeaderSet[idx] = ''
          retVal = true
          for (var idx2 = idx; idx2 < localHeaderSet.length; idx2++) {
            if (localHeaderSet[idx2].toUpperCase().startsWith(EMAIL_CONTENT_TRANSFER_MARKER)) {
              if (localHeaderSet[idx2].toUpperCase().indexOf(EMAIL_BINARY_CONTENT_MARKER) > 0) {
                // Replace the header with an empty string.
                localHeaderSet[idx2] = ''
                // true if we removed the content-type and content-transfer encoding headers
                break
              }
            }
          }
          break;
        }
      }
    }
    // We update the headerset passed by the calling function
    usrObject.headerArray = localHeaderSet
    // false if the headers were untouched
    return retVal
  }

  /**
   * Checks whether first character in string is a control character, a white space or a colon.
   *
   * @param {string} rawString to be checked
   * 
   * @return {boolean} true if first character is CTL, WSP or colon. false if not.
   */
  function isFirstCharAControlCharASpaceOrColon(rawString) {
    // field       =  field-name ":" [ field-body ] CRLF   
    // field-name  =  1*<any CHAR, excluding CTLs, SPACE, and ":">
    // field-body  =  field-body-contents
    //                [CRLF LWSP-char field-body]
    // field-body-contents =
    // <the ASCII characters making up the field-body, as
    //   defined in the following sections, and consisting
    //   of combinations of atom, quoted-string, and
    //   specials tokens, or else consisting of texts>
    // CTLs 0-31 and space character 
    // return true if the rawstring is empty
    if (!rawString) return true
    if (rawString[0] <= 0x32) {
      return true
    }
    if (rawString[0] == ':') {
      return true
    }
    if (rawString[0] == 0x7F) {
      return true
    }
    return false
  }

  /**
   * RFC-822 Based string parsing of email data, that returns significant details of the email like the offset
   * where the start of the email body is found and splits the email-string into an array of headers
   *
   * @param {string} rawEmailDataAsString the email content, headers and body, as string
   * 
   * @return {object} Returns an object with the following members :
   *      parsedHeaders - email headers now written as an array.
   *      bodyOffset - character index where the "body" of the email is found. (-1 if not found)
   *      contentTypeLine - content-type header name and value
   *      mimeVersionLine - mime version header and value
   */
  function getHeadersAndSignificantLines(rawEmailDataAsString) {
    var parsedHeadersArray = []
    var contentTypeLine = ''
    var mimeVersionLine = ''
    var marker = "\r\n"
    var tmpBodyOffset = 0
    var bodyOffset = -1

    var localHeaderArray = rawEmailDataAsString.split(marker)
    if (localHeaderArray.length <= 1) {
      marker = "\n"
      localHeaderArray = rawEmailDataAsString.split(marker)
    }
    //now we expect the string to have been split, but we need to acquire the correct header.
    var tmpHdrNameAndValue = ''
    var wasContentTypeFound = false
    var wasMimeVersionFound = false
    for (var idx = 0; idx < localHeaderArray.length; idx++) {
      if (!isFirstCharAControlCharASpaceOrColon(localHeaderArray[idx])) {
        // if the first character in the line is not a space, a Control Char, or colon, we mark it as a new header.
        // we first check if there is a colon 
        // so it starts with a valid character, is there a colon in it?
        var indexOfColon = localHeaderArray[idx].indexOf(':')
        // if the line has no colon, and does not start with CTL, space or colon, then it is most likely body
        if (indexOfColon < 0) {
          if (tmpHdrNameAndValue) {
            // we now save the value of tmpHdrNameAndValue
            parsedHeadersArray.push(tmpHdrNameAndValue)
            tmpBodyOffset += tmpHdrNameAndValue.length
          }
          bodyOffset = tmpBodyOffset
          break;
        }
        // A colon was found, and that our temporary header name and value is non-empty. 
        // Then this line is a new header.
        if (tmpHdrNameAndValue) {
          // we now save the value of tmpHdrNameAndValue
          parsedHeadersArray.push(tmpHdrNameAndValue)
          tmpBodyOffset += tmpHdrNameAndValue.length
          if (!wasContentTypeFound && tmpHdrNameAndValue.toUpperCase().startsWith(EMAIL_CONTENT_TYPE_MARKER)) {
            contentTypeLine = tmpHdrNameAndValue
            wasContentTypeFound = true
          }
          if (!wasMimeVersionFound && tmpHdrNameAndValue.toUpperCase().startsWith(EMAIL_MIME_VERSION_MARKER)) {
            mimeVersionLine = tmpHdrNameAndValue
            wasMimeVersionFound = true
          }
          tmpHdrNameAndValue = localHeaderArray[idx] + marker
        } else {
          tmpHdrNameAndValue += localHeaderArray[idx] + marker
        }
      } else {
        // so the first character is a space, colon or control character. We need to check if for null first.
        // So is the next line an empty line???
        if (localHeaderArray[idx] == '') {
          if (tmpHdrNameAndValue) {
            // we now save the value of tmpHdrNameAndValue
            parsedHeadersArray.push(tmpHdrNameAndValue)
            tmpBodyOffset += tmpHdrNameAndValue.length
          }
          // this is already the body, let us get the total body offset.
          bodyOffset = tmpBodyOffset
          break;
        } else {
          // if it is not an empty line, we add it to our parsed headers.
          tmpHdrNameAndValue += localHeaderArray[idx] + marker
        }
      }
    }
    return {
      "parsedHeaders": parsedHeadersArray,
      "bodyOffset": bodyOffset,
      "contentTypeLine": contentTypeLine,
      "mimeVersionLine": mimeVersionLine
    }
  }

  /**
   * Builds an eml-string, using the body of the original mimecontent and the full header set provided. 
   *
   * @param {string} encodedMimeContent the base64 encoded mimecontent provided by the EWS server
   * @param {string} emailHeaders the full header set, expressed as one string, provided by the EWS server
   * 
   * @return {string} returns an eml file as a string.
   */
  function buildEMLFileFromMimeContentAndFullHeaders(encodedMimeContent, emailHeaders) {
    // first let us decode data, if it is needed.
    var decodedMimeContent = atob(encodedMimeContent)
    //var decodedMimeContent =Buffer.from(encodedMimeContent, 'base64').toString()
    var decodedHeaders = emailHeaders.toString()
    var bodyOffset = 0

    var parsedHeadersFromHeaderSet = []
    var contentTypeLineFromHeaderSet = ''
    var mimeVersionLineFromHeaderSet = ''
    var contentTypeLineFromMimeContent = ''
    var mimeVersionLineFromMimeContent = ''

    // We parse the complete headers set first and remember the contentTypeLine,  mimeVersionLine. Body offset of headers is not used.
    var parsedEmailObject = getHeadersAndSignificantLines(decodedHeaders)
    contentTypeLineFromHeaderSet = parsedEmailObject.contentTypeLine
    parsedHeadersFromHeaderSet = parsedEmailObject.parsedHeaders
    mimeVersionLineFromHeaderSet = parsedEmailObject.mimeVersionLine

    parsedEmailObject = getHeadersAndSignificantLines(decodedMimeContent)
    contentTypeLineFromMimeContent = parsedEmailObject.contentTypeLine
    bodyOffset = parsedEmailObject.bodyOffset
    mimeVersionLineFromMimeContent = parsedEmailObject.mimeVersionLine

    // clean up the header set of unwanted winmail.dat headers if necessary
    if (contentTypeLineFromHeaderSet.toUpperCase() != contentTypeLineFromMimeContent.toUpperCase()) {
      //we create a temporary object for our arrays.
      var headerSetObject = { "headerArray": parsedHeadersFromHeaderSet }
      removeWinMailDatHeader(headerSetObject)
      // now we add the right contenttype
      parsedHeadersFromHeaderSet.push(contentTypeLineFromMimeContent)
    }
    // finally we add the mime version
    if (!mimeVersionLineFromHeaderSet && mimeVersionLineFromMimeContent) {
      parsedHeadersFromHeaderSet.push(mimeVersionLineFromMimeContent)
    }
    // Let us rebuild the eml file by joining all full headers and then adding the "body" part of the mime content.
    return (parsedHeadersFromHeaderSet.join('') + decodedMimeContent.substring(bodyOffset))
  }


  /**
   * Opens a dialog as an iframe to display the success message.
   *
   * @param {string} okLabelTxt the text for the "OK" button to
   * @param {string} rawUsrSuccessMessage the success message to be displayed.
   * @param {integer} usrTimeoutInMS the duration as to how long the message will be displayed in milliseconds.
   * 
   * @return {void} 
   */
  function openDialogAsIframe(okLabelTxt, rawUsrSuccessMessage, usrTimeoutInMS) {
    // hide the old notification.
    //IMPORTANT: IFrame mode only works in Online (Web) clients. Desktop clients (Windows, IOS, Mac) always display as a pop-up inside of Office apps. 
    var windowFullPath = window.location.href
    var finalURL = ''
    // https://od-pb-898-hybrid.kmsat.internal.knowbe4.com/phishalertonline/CC9227FC0491EC658BF64F30B437E68C/AppRead/Home/Home.html?et=
    if (windowFullPath) {
      var indexOfAppread = windowFullPath.toUpperCase().indexOf(PABV1_KB4_URLMARKER)
      if (indexOfAppread > -1) {
        finalURL = windowFullPath.substring(0, indexOfAppread);
        // force the use of a blank space if rawUsrMessage is null.
        var usrSuccessMessage = !rawUsrSuccessMessage ? ' ' : rawUsrSuccessMessage;
        //var userAgentAndPlatform="[" + ua + "]["+ plat + "]"
        var encodedUsrMsg = window.btoa(unescape(encodeURIComponent(usrSuccessMessage)))
        var encodedOkLabel = window.btoa(unescape(encodeURIComponent(okLabelTxt)))
        displayDurationInMs = usrTimeoutInMS
        var finalPath = finalURL + "Static/NotificationDialog.html" + "?okButtonLabel=" + encodedOkLabel + "&usrConfMg=" + encodedUsrMsg
        // hide the notification bar first, then display the async dialog.
        $('#notification-message').hide()
        Office.context.ui.displayDialogAsync(finalPath, { height: 20, width: 50, displayInIframe: true }, dialogCallback)
      }
    } else {
      // we force a throw, here to make the upper function handle it.
      throw 'System does not support OpenDialogAPI.'
    }
  }

  /**
   * callback function when the dialogcallback is executed.
   *
   * @param {object} asyncResult Blob of data containing the resulf ot displayDialogAsync()
   * 
   * @return {void} 
   */
  function dialogCallback(asyncResult) {
    // let us raise the needed flags.
    isToDeleteEmail = true
    isCloseDialog = true
    notifDialog = asyncResult.value
    notifDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage)
    notifDialog.addEventHandler(Office.EventType.DialogEventReceived, handleDialogEvent)
    //otherwise close notification after user defined timeout
    setTimeout(closeAndDelete, displayDurationInMs)
    return
  }

  /**
   * Function invoked whenever the success-dialog closes after a user executes an action.
   *
   * @param {object} childMessage unused data blob from the dialog message.
   * 
   * @return {void} 
   */
  function processMessage(childMessage) {
    // Regardless of message, we need to delete the email.
    closeAndDelete()
    return
  }

  /**
   * Function invoked whenever the success-dialog closes after time expires
   * 
   * @return {void} 
   */
  function handleDialogEvent() {
    // the dialog was manually closed.
    isCloseDialog = false
    closeAndDelete()
    return
  }

  /**
   * Closes the taskpane and then deletes the message being displayed.
   * 
   * @return {void} 
   */
  function closeAndDelete() {
    hideLoader()
    if (isToDeleteEmail) {
      isToDeleteEmail = false
      if (isCloseDialog) {
        notifDialog.close()
      }
      deleteEmail(true)
      notifDialog = null
      isToDeleteEmail = false
      isCloseDialog = false
    }
    return
  }

  /**
   * Gets a token to access the message, and then deletes the message being displayed.
   * 
   * @param {boolean} isFromAsyncDialog true if originating call is an asyncdialog.
   * 
   * @return {void} 
   */
  function getChangeKeyAndDeleteEmail(isFromAsyncDialog) {
    // Regardless of results, we need to delete the email.
    var mailbox = Office.context.mailbox
    var item = Office.cast.item.toItemRead(Office.context.mailbox.item)
    mailbox.makeEwsRequestAsync(getItemDataRequest(item.itemId, ['Subject']), getItemToDeleteCallback, { isFromAsyncDialog: isFromAsyncDialog })
    return
  }

})()
