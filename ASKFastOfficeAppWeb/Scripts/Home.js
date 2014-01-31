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

            //create toggle effect for tabs
            $('#myTabs a').click(function (e) {
                e.preventDefault()
                $(this).tab('show')
            });
            $('#generate').click(genereateReport);
        });
    };

    //enable extras if text message (apart from twitter) is selected
    function toggleAdapterTypes() {
        if ($('#xmpp').is(":checked") || $('#sms').is(":checked") || $('#mail').is(":checked")) {
            $('#extras').show();
            $('#extrasHeader').show();
            $('#broadcastMessage').html('<strong>Step 7: </strong> Broadcast your message.');
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
            $('#broadcastMessage').html('<strong>Step 6: </strong> Broadcast your message.');
            $('#extrasHeader').hide();
            $('#extras').hide();
        }
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
            var broadcastNode = new Object();
            var adapterList = new Object();
            adapterList["EMAIL"] = "bb3fe5e0-84de-11e3-998f-00007f000001";
            broadcastNode["adapterList"] = adapterList;
            broadcastNode["senderName"] = $('#senderId').val();
            broadcastNode["message"] = $('#message').val();
            broadcastNode["retryMethod"] = 'MANUAL';
            broadcastNode["broadcastName"] = 'Broadcast from ASK-Fast Excel App';
            broadcastNode["emailSubject"] = $('#subject').val();
            broadcastNode["language"] = $('#language').val();
            json["broadcast"] = broadcastNode;
            app.showNotification("Sending your request...", "");
            $.ajax({
                cache: false,
                crossDomain: true,
                contentType: 'application/json; charset=utf-8',
                url: '/App/Handler1.ashx?' + 'questionType=' + $('#questionType').val() + appendHeaders(),
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

    //returns a csv format string from the excel range selected
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

    //returns a string value of all the header query parameters added
    function appendHeaders() {
        return "&firstName=" + encodeURIComponent(firstNameHeader) +
            "&lastName=" + encodeURIComponent(lastNameHeaderHeader) +
            "&phoneHeader=" + encodeURIComponent(phoneHeader) +
            "&emailHeader=" + encodeURIComponent(emailHeader) +
            "&xmppHeader=" + encodeURIComponent(xmppHeader) +
            "&facebookHeader=" + encodeURIComponent(facebookHeader) +
            "&twitterHeader=" + encodeURIComponent(twitterHeader);
    }

    //generates the report on the excel sheet for the reponse to the questions seen
    function genereateReport() {
        $.ajax({
            contentType: 'application/json; charset=utf-8',
            url: '/App/Handler1.ashx?' + 'questionType=' + $('#questionType').val() + appendHeaders(),
            type: 'GET',
            dataType: 'json'
        }).success(function (response) {
            writeReportOnSheet(response);
            app.showNotification("Success", response.statusText);
            console.log("Success", response.statusText);
        }).error(function (response) {
            app.showNotification("Error", response.statusText);
            console.log("Error", response.statusText);
        });
    }

    //function to write report on the sheet
    function writeReportOnSheet(response) {
        if (response != null && response.length != 0) {
            var data = new Object();
            var rowCounter = 0;
            data[rowCounter] = ["Timestamp", "Question", "Responder", "Response"];
            for (; rowCounter < response.length; rowCounter++) {
                var questionResponse = response[rowCounter];
                var rowData = new Object();
                rowData[0] = questionResponse["timestamp"];
                var questionMap = questionResponse["clipboardMap"];
                rowData[1] = "";
                if (questionMap["question"] != null) {
                    var question = JSON.parse((questionMap["question"]));
                    var questionText = question["question_text"];
                    if (questionText != null) {
                        questionText = questionText.replace('text://', '');
                    }
                    rowData[1] = decodeURIComponent(questionText);
                }
                rowData[2] = questionMap["responder"];
                rowData[3] = questionMap["answer_text"];
                data[rowCounter + 1] = rowData;
            }
            Office.context.document.setSelectedDataAsync(data, { coercionType: Office.CoercionType.Matrix });
        }
        else {
            app.showNotification("Info", "No reports found.")
        }
    }
})();