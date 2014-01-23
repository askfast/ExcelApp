/// <reference path="../App.js" />

(function () {
    "use strict";
    var phoneHeader = "Phone";
    var firstNameHeader = "First Name";
    var lastNameHeaderHeader = "Last Name";
    var emailHeader = "Email";
    var xmppHeader = "XMPP";
    var facebookHeader = "Facebook";
    var twitterHeader = "Twitter";

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
        //fix a binding element to the addresses selected
        Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Matrix, { id: 'addresses' },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    result.value.getDataAsync(getDataFromBinding);
                }
            });
    };

    function getDataFromBinding(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            var csvData = getCSVFromSelection(result.value);
            var json = new Object();
            json["csvStream"] = csvData;
            json["senderName"] = $('#senderId').val();
            json["message"] = $('#message').val();
            json["retryMethod"] = 'MANUAL';
            json["broadcastName"] = 'Broadcast from ASK-Fast Excel App';
            json["emailSubject"] = $('#subject').val();
            json["language"] = $('#language').val();
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
    }

    function getCSVFromSelection(result) {
        var resultCSV = '';
        for (var rowCount = 0; rowCount < result.length; rowCount++) {
            for (var columnCount = 0; columnCount < result[rowCount].length; columnCount++) {
                resultCSV += result[rowCount][columnCount];
                if (columnCount != result[rowCount].length - 1) {
                    resultCSV += ",";
                }
            }
            if (rowCount != result.length - 1) {
                resultCSV += "\n";
            }
        }
        //check if a single cell is selected i.e one row and one column range
        if (result.length == 1 && result[0].length == 1) {
            
        }
        return resultCSV;
    }

    //converts the selected excel data to csv format
    function enableSendButton(result) {
        if ($('#message').val() != "") {
            $('#send').removeAttr("disabled");
        }
    };

    //get all the channels selected
    function getChannelsChecked() {
        var resultChannels = new Object();
        var channelCounter = 0;
        if ($('#xmpp').is(":checked")) {
            resultChannels[channelCounter++] = $('#xmpp').val();
        }
        if ($('#mail').is(":checked")) {
            resultChannels[channelCounter++] = $('#mail').val();
        }
        if ($('#broadsoft').is(":checked")) {
            resultChannels[channelCounter++] = $('#broadsoft').val();
        }
        if ($('#sms').is(":checked")) {
            resultChannels[channelCounter++] = $('#sms').val();
        }
        if ($('#twitter').is(":checked")) {
            resultChannels[channelCounter++] = $('#twitter').val();
        }
        return resultChannels;
    }
})();