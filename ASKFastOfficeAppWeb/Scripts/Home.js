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
            $('#broadcastMessage').html('<strong>Step 6: </strong> Broadcast your message.');
            if ($('#xmpp').is(":checked") || $('#sms').is(":checked")) {
                $('#senderIdRow').show();
                $('#subjectRow').hide();
            }
            if ($('#mail').is(":checked")) {
                $('#senderIdRow').show();
                $('#subjectRow').show();
            }
        }
        else {
            $('#broadcastMessage').html('<strong>Step 5: </strong> Broadcast your message.');
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
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var csvData = getCSVFromSelection(result);
                    var json = new Object();
                    json["csvStream"] = csvData;
                    json["senderName"] = $('#senderId').val();
                    json["message"] = $('#message').val();
                    json["retryMethod"] = 'MANUAL';
                    json["broadcastName"] = 'Broadcast from ASK-Fast Excel App';
                    json["emailSubject"] = $('#subject').val();
                    json["language"] = $('#language').val();

                    
                    ////build address payload. Refer to https://docs.google.com/a/ask-cs.com/document/d/1J7ceZAy39ZZMc4k8CGivNbkz94zDskGcklriUu0bVzs/edit#bookmark=kix.jxex4x2y5zfl
                    ////each address must be an element of {}
                    //var addressCollection = new Object();
                    //var addressNames = result.value.split("\n");
                    //for (var addressCount = 0; addressCount < addressNames.length && addressNames[addressCount] != "" ; addressCount++) {
                    //    var singleAddressNode = new Object();
                    //    var addressWithName = addressNames[addressCount].split("\t");
                    //    if (addressWithName.length != 0 && addressWithName[0] != "") {
                    //        singleAddressNode["address"] = addressWithName[0];
                    //        addressCollection[addressCount] = singleAddressNode;
                    //    }
                    //}
                    //var addresses = new Object();
                    //params["url"] = "http://askfastmarket1.appspot.com/resource/question/" + encodeURIComponent($('#message').val());
                    //params["addressMap"] = addressMap;
                    //params["publicKey"] = publicKey;
                    //params["privateKey"] = privateKey;
                    //params["senderName"] = $('#senderId').val();
                    //json["params"] = params;
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

    function getCSVFromSelection(result) {
        if (result.status == 'succeeded') {
            var resultCSV = '';
            for (var rowCount = 0; rowCount < result.value.length; rowCount++) {
                for (var columnCount = 0; columnCount < result.value[rowCount].length; columnCount++) {
                    resultCSV += result.value[rowCount][columnCount];
                    if (columnCount != result.value[rowCount].length - 1) {
                        resultCSV += ",";
                    }
                }
                if (rowCount != result.value.length - 1) {
                    resultCSV += "\n";
                }
            }
            return resultCSV;
        }
    }

    //converts the selected excel data to csv format
    function enableSendButton(result) {
        if ($('#message').val() != "") {
            $('#send').removeAttr("disabled");
        }
    };
})();