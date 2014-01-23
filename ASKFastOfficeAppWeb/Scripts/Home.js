/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            if ($('#send').val() == "") {
                $('#send').attr("disabled", "disabled");
            }
            $('#message').keyup(enableSendButton);
            $('#get-data-from-selection').click(getDataFromSelection);
            $('#send').click(getDataFromSelection);
            //disable the extras div
            $('#extrasHeader').hide();
            $('#extras').hide();
            //event listners for adapters toggle 
            toggleAdapterTypes();
            $('#xmpp').click(toggleAdapterTypes);
            $('#mail').click(toggleAdapterTypes);
            $('#broadsoft').click(toggleAdapterTypes);
            $('#sms').click(toggleAdapterTypes);
            $('#twitter').click(toggleAdapterTypes);
        });
    };

    //enable extras if text message (apart from twitter) is selected
    function toggleAdapterTypes() {
        if ($('#xmpp').is(":checked") || $('#sms').is(":checked") || $('#mail').is(":checked")) {
            $('#extras').show();
            $('#extrasHeader').show();
            $('#broadcastMessage').html('<strong>Step 5: </strong> Broadcast your message.');
            if ($('#xmpp').is(":checked") || $('#sms').is(":checked")) {
                $('#senderId').show();
                $('#subject').hide();
            }
            if ($('#mail').is(":checked")) {
                $('#senderId').show();
                $('#subject').show();
            }
        }
        else {
            $('#broadcastMessage').html('<strong>Step 4: </strong> Broadcast your message.');
            $('#extrasHeader').hide();
            $('#extras').hide();
        }
    }

    //
    //showResult if outbound call was successful
    function showResult(response) {
        app.showNotification("Response:", response);
    }

    //showResult if outbound call was successful
    function showError(response) {
        app.showNotification("Error:", response.statusText);
    }

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var json = new Object();
                    json["method"] = "outboundCallWithMap";
                    var params = new Object();
                    var adapterId = null;
                    var publicKey = "";
                    var privateKey = "";
                    switch ($('#adapterType').val())
                    {
                        case "mail":
                            adapterId = "cc3b0c20-2ffd-11e3-a94b-00007f000001";
                            publicKey = "oneline@askfast.com";
                            break;
                        case "xmpp":
                            adapterId = "f3ccf8f0-6b92-11e2-b94c-00007f000001";
                            publicKey = "5368dbd0-058f-11e3-a6c9-060dc6d9dd94";
                            break;
                        case "broadsoft":
                            adapterId = "151b05e0-25f3-11e3-a6c7-00007f000001";
                            break;
                        case "sms":
                            adapterId = "3c9e7300-0e4b-11e3-837b-00007f000001";
                            break;
                        case "twitter":
                            adapterId = "e8f42228-13a9-406f-b991-748d1a61504d";
                            break;
                    }
                    if (adapterId == null) {
                        params["adapterID"] = $('#adapterType').val().toUpperCase;
                    }
                    else {
                        params["adapterID"] = adapterId;
                    }
                    var addressMap = new Object();
                    var addressNames = result.value.split("\n");
                    for (var addressCount = 0; addressCount < addressNames.length && addressNames[addressCount] != "" ; addressCount++) {
                        var addressWithName = addressNames[addressCount].split("\t");
                        if (addressWithName.length != 0 && addressWithName[0] != "") {
                            if (addressWithName.length > 1) {
                                addressMap[addressWithName[0]] = addressWithName[1];
                            }
                            else {
                                addressMap[addressWithName[0]] = "";
                            }
                        }
                    }
                    params["url"] = "http://askfastmarket1.appspot.com/resource/question/" + encodeURIComponent($('#message').val());
                    params["addressMap"] = addressMap;
                    params["publicKey"] = publicKey;
                    params["privateKey"] = privateKey;
                    params["senderName"] = $('#senderId').val();
                    json["params"] = params;
                    app.showNotification("Sending your request...", "");

                    $.ajax({
                        cache: false,
                        crossDomain: true,
                        contentType: 'application/json; charset=utf-8',
                        url: '/App/Handler1.ashx',
                        type: 'POST',
                        dataType: 'json',
                        jsonpCallback: function (response) {
                            app.showNotification("arguments", Array.prototype.join.call(arguments, ' '));
                        },
                        data: JSON.stringify(json)
                    }).success(function (response) {
                        app.showNotification("Success", response.statusText);
                        console.log("Success", response.statusText);
                    }).error(function (response) {
                        app.showNotification("Error", response.statusText);
                        console.log("Error", response.statusText);
                    });
                } else {
                    app.showNotification('Error:', result.error.message);
                }
                return null;
            }
        );
    };

    //enables the send button, only when some message is entered
    function enableSendButton() {
        if ($('#message').val() != "") {
            $('#send').removeAttr("disabled");
        }
    };
})();