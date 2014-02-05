/// <reference path="../App.js" />

(function () {
    "use strict";
    //default headers
    var smsHeader = "Mobile";
    var fixedLineHeader = "Fixed Line";
    var firstNameHeader = "First Name";
    var lastNameHeaderHeader = "Last Name";
    var emailHeader = "Email";
    var xmppHeader = "XMPP";
    var facebookHeader = "Facebook";
    var twitterHeader = "Twitter";
    //header mappings to adapters
    var adapterMappings = {};
    var X_SESSIONID = "";
    var lastTabSelected = "#homeTab"; //default it to hte active tab

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            enableSendButton();
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
            $('#login').click(doLogin);
            enableLoginButton();
            $('#username').keyup(enableLoginButton);
            $('#password').keyup(enableLoginButton);

            //create toggle effect for tabs
            $('#myTabs a').click(function (e) {
                e.preventDefault();
                $(this).tab('show');
                if ($(e.target).attr('href') != '#loginTab') {
                    lastTabSelected = $(e.target).attr('href');
                }
            });
            $('#generate').click(genereateReport);
        });
    };

    //perform login
    function doLogin() {
        if ($('#username').val() != null && $('#password').val() != null) {
            app.showNotification("Signing in into " + $('#username').val(), "");
            $.ajax({
                cache: false,
                crossDomain: true,
                contentType: 'application/json; charset=utf-8',
                url: '/App/Handler1.ashx/login?username=' + $('#username').val() + "&password=" + CryptoJS.MD5($('#password').val()).toString(),
                type: 'GET',
                dataType: 'json'
            }).success(function (response) {
                X_SESSIONID = response["X-SESSION_ID"];
                app.showNotification("Success", "Login successful");
                console.log("Success", response.statusText);
                if (lastTabSelected) {
                    $('a[href="' + lastTabSelected + '"]').tab('show');
                }
            }).error(function (response) {
                app.showNotification("Error", response.responseText);
                console.log("Error", response.statusText);
            });
        }
    }
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
        if (X_SESSIONID != null && X_SESSIONID != "") {
            updateHeaders();
            //fix a binding element to the addresses selected
            Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Matrix, { id: 'addresses' },
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        result.value.getDataAsync(getDataFromBinding);
                    }
                    else {
                        app.showNotification(result.error.name, result.error.message)
                    }
                });
        }
        else {
            app.showNotification("Authentication error", "Please login..");
            $('a[href="#loginTab"]').tab('show');
        }
    };

    function getDataFromBinding(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            var csvData = getCSVFromSelection(result.value);
            var json = new Object();
            json["csvStream"] = csvData;
            var broadcastNode = new Object();
            var adapterList = new Object();
            broadcastNode["senderName"] = $('#senderId').val() != null ?
                $('#senderId').val() : "ASKFast4Excel";
            broadcastNode["message"] = $('#message').val();
            broadcastNode["retryMethod"] = 'MANUAL';
            broadcastNode["broadcastName"] = 'Broadcast from ASK-Fast Excel App';
            broadcastNode["emailSubject"] = $('#subject').val() != null ?
                $('#subject').val() : 'Broadcast from ASK-Fast Excel App';
            broadcastNode["language"] = $('#language').val();
            json["broadcast"] = broadcastNode;
            app.showNotification("Sending your request...", "");
            $.ajax({
                cache: false,
                crossDomain: true,
                contentType: 'application/json; charset=utf-8',
                beforeSend: function (request) {
                    request.setRequestHeader("X-SESSION_ID", X_SESSIONID);
                },
                url: '/App/Handler1.ashx?' + 'questionType=' + $('#questionType').val() + appendHeaders(),
                type: 'POST',
                dataType: 'json',
                jsonpCallback: function (response) {
                    app.showNotification("arguments", Array.prototype.join.call(arguments, ' '));
                },
                data: JSON.stringify(json)
            }).success(function (response) {
                showFailures(response);
            }).error(function (response) {
                app.showNotification("Error", response.responseText);
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
                //check if the column is ignored
                if (!shouldColumnBeIgnored(result[0][columnCount])) {
                    if (columnCount != 0) {
                        resultCSV += ",";
                    }
                    resultCSV += result[rowCount][columnCount];
                }
            }
            if (rowCount != result.length - 1) {
                resultCSV += "\n";
            }
        }
        return resultCSV;
    }

    function enableSendButton(result) {
        if ($('#message').val() != "") {
            $('#send').removeAttr("disabled");
        }
    };

    function enableLoginButton(result) {
        if ($('#username').val() != "" && $('#password').val() != "") {
            $('#login').removeAttr("disabled");
        }
    };

    //show failures seen while broadcasting
    function showFailures(response) {
        if (response['broadcast'] && response['broadcast']['addresses']) {
            var failureMessage = "";
            for (var selectedMedium in response['broadcast']['addresses']) {
                for (var messageIndex in response['broadcast']['addresses'][selectedMedium]) {
                    var message = response['broadcast']['addresses'][selectedMedium][messageIndex]['responseMessage'];
                    if (message) {
                        failureMessage += message + '\n';
                    }
                }
            }
            if (failureMessage != "") {
                app.showNotification("Details", failureMessage);
            }
        }
    }

    //get all the channels selected
    function getChannelsChecked() {
        var resultChannels = new Array();
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

    //update the default columnHeaders if changes in settings tab
    function updateHeaders() {
        //update the headers
        smsHeader = $('#smsHeader').val() != smsHeader ? $('#smsHeader').val() : smsHeader;
        fixedLineHeader = $('#fixedLineHeader').val() != fixedLineHeader ? $('#fixedLineHeader').val() : fixedLineHeader;
        firstNameHeader = $('#firstNameHeader').val() != firstNameHeader ? $('#firstNameHeader').val() : firstNameHeader;
        lastNameHeaderHeader = $('#lastNameHeader').val() != lastNameHeaderHeader ? $('#lastNameHeader').val() : lastNameHeaderHeader;
        emailHeader = $('#emailHeader').val() != emailHeader ? $('#emailHeader').val() : emailHeader;
        xmppHeader = $('#xmppHeader').val() != xmppHeader ? $('#xmppHeader').val() : xmppHeader;
        twitterHeader = $('#twitterHeader').val() != twitterHeader ? $('#twitterHeader').val() : twitterHeader;
        //update the adapterMappings
        adapterMappings[fixedLineHeader] = "broadsoft";
        adapterMappings[smsHeader] = "sms";
        adapterMappings[xmppHeader] = "xmpp";
        adapterMappings[emailHeader] = "mail";
        adapterMappings[twitterHeader] = "twitter";
    }

    //returns true or false based on the mediums selected in the checkbox: adapterTypes
    function shouldColumnBeIgnored(header) {
        //if column is not part of adapterMappings include it. eg firstName, lastName
        if (adapterMappings[header] == null) {
            return false;
        }
        else if (adapterMappings[header] != null
            && $.inArray(adapterMappings[header], getChannelsChecked()) != -1) {
            return false;
        }
        return true;
    }

    //returns a string value of all the header query parameters added
    function appendHeaders() {
        return "&firstName=" + encodeURIComponent(firstNameHeader) +
            "&lastName=" + encodeURIComponent(lastNameHeaderHeader) +
            "&smsHeader=" + encodeURIComponent(smsHeader) +
            "&callHeader=" + encodeURIComponent(fixedLineHeader) +
            "&emailHeader=" + encodeURIComponent(emailHeader) +
            "&xmppHeader=" + encodeURIComponent(xmppHeader) +
            "&facebookHeader=" + encodeURIComponent(facebookHeader) +
            "&twitterHeader=" + encodeURIComponent(twitterHeader);
    }

    //generates the report on the excel sheet for the reponse to the questions seen
    function genereateReport() {
        if (X_SESSIONID != null && X_SESSIONID != "") {
            app.showNotification("Generating report..", "");
            $.ajax({
                contentType: 'application/json; charset=utf-8',
                beforeSend: function (request) {
                    request.setRequestHeader("X-SESSION_ID", X_SESSIONID);
                },
                url: '/App/Handler1.ashx',
                type: 'GET',
                dataType: 'json'
            }).success(function (response) {
                writeReportOnSheet(response);
                app.showNotification("Success", response.statusText);
                console.log("Success", response.statusText);
            }).error(function (response) {
                app.showNotification("Error", response.responseText);
                console.log("Error", response.statusText);
            });
        }
        else {
            app.showNotification("Authentication error", "Please login..");
            $('a[href="#loginTab"]').tab('show');
        }
    }

    //function to write report on the sheet
    function writeReportOnSheet(response) {
        if (response != null && response.length != 0) {
            var data = new Object();
            var rowCounter = 0;
            data[rowCounter] = ["Timestamp", "Question Type", "Question", "Responder", "Responder Name",
                "Status", "Response"];
            for (; rowCounter < response.length; rowCounter++) {
                var questionResponse = response[rowCounter];
                var rowData = new Object();
                //initialize with empty values
                for (var columnCounter = 0; columnCounter < data[0].length; columnCounter++) {
                    rowData[columnCounter] = "";
                }
                rowData[0] = getStringDateFromMilliseconds(questionResponse["timestamp"]);
                var questionMap = questionResponse["clipboardMap"];
                if (questionMap["question"] != null) {
                    var question = JSON.parse((questionMap["question"]));
                    rowData[1] = question["type"]
                    var questionText = question["question_text"];
                    if (questionText != null) {
                        questionText = questionText.replace('text://', '');
                    }
                    rowData[2] = decodeURIComponent(questionText);
                }
                rowData[3] = questionMap["responder"];
                rowData[4] = questionMap["responder_name"];
                rowData[5] = questionMap["status"];
                rowData[6] = questionMap["answer_text"];
                data[rowCounter + 1] = rowData;
            }
            Office.context.document.setSelectedDataAsync(data, { coercionType: Office.CoercionType.Matrix });
        }
        else {
            app.showNotification("Info", "No reports found.")
        }
    }

    function getStringDateFromMilliseconds(milliseconds) {
        var date = new Date(milliseconds);
        return date.toDateString() + " " + date.toLocaleTimeString();
    }
})();