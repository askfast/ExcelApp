/// <reference path="../App.js" />

(function () {
    "use strict";
    //default headers
    var smsHeader = "Mobile";
    var fixedLineHeader = "Fixed Line";
    var firstNameHeader = "First Name";
    var lastNameHeader = "Last Name";
    var emailHeader = "Email";
    var xmppHeader = "XMPP";
    var facebookHeader = "Facebook";
    var twitterHeader = "Twitter";
    //header mappings to adapters
    var headerMappings = {};
    var X_SESSIONID = "";
    var lastTabSelected = "#homeTab"; //default it to hte active tab

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            //login if username and password is enabled already
            if (supports_html5_storage()) {
                $('#username').val(localStorage.getItem("username"));
                $('#password').val(localStorage.getItem("password"));
                if ($('#username').val() != null && $('#password').val() != null
                    && $('#autoLogon').is(":checked")) {
                    doLogin();
                }
            }
            loadDataFromLocalStorage();
            enableSendButton();
            $('#message').keyup(enableSendButton);
            $('#send').click(getDataFromSelection);
            //disable the extras div
            $('#extrasHeader').hide();
            $('#extras').hide();
            //event listners for adapters toggle 
            toggleAdapterTypes();
            toggleMatchingQuestionText();
            $('#xmpp').click(toggleAdapterTypes);
            $('#mail').click(toggleAdapterTypes);
            $('#broadsoft').click(toggleAdapterTypes);
            $('#sms').click(toggleAdapterTypes);
            $('#twitter').click(toggleAdapterTypes);
            $('#matchQuestionCheckBox').click(toggleMatchingQuestionText);
            $('#login').click(doLogin);
            enableLoginButton();
            $('#username').keyup(enableLoginButton);
            $('#password').keyup(enableLoginButton);
            $(".nexttab").click(function () {
                $('a[href="#settingsTab"]').tab('show');
            });

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
            //store the page values in local storage
            if (supports_html5_storage()) {
                localStorage.setItem("username", $('#username').val());
                localStorage.setItem("password", $('#password').val());
            }

            app.showNotification("Signing in into " + $('#username').val(), "");
            $.ajax({
                cache: false,
                crossDomain: true,
                contentType: 'application/json; charset=utf-8',
                url: './ASKFastRequestHandler.ashx/login?username=' + $('#username').val() + "&password=" + CryptoJS.MD5($('#password').val()).toString(),
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
            $('#broadcastMessage').html('<strong>Step 7: </strong> Send your message.');
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
            storeDataInLocalStorage();
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
            if ($('#questionType').val() == 'closed') {
                broadcastNode["answers"] = ["Yes", "No"];
            }
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
                url: './ASKFastRequestHandler.ashx?' + 'questionType=' + $('#questionType').val() + appendHeaders(),
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
            else {
                app.showNotification("Success", "Message sent successfully!");
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
        //ugly hack: including the firstName and the lastName headers
        resultChannels[channelCounter++] = "firstName";
        resultChannels[channelCounter++] = "lastName";
        return resultChannels;
    }

    //update the default columnHeaders if changes in settings tab
    function updateHeaders() {
        //update the headers
        smsHeader = $('#smsHeader').val() != smsHeader ? $('#smsHeader').val() : smsHeader;
        fixedLineHeader = $('#fixedLineHeader').val() != fixedLineHeader ? $('#fixedLineHeader').val() : fixedLineHeader;
        firstNameHeader = $('#firstNameHeader').val() != firstNameHeader ? $('#firstNameHeader').val() : firstNameHeader;
        lastNameHeader = $('#lastNameHeader').val() != lastNameHeader ? $('#lastNameHeader').val() : lastNameHeader;
        emailHeader = $('#emailHeader').val() != emailHeader ? $('#emailHeader').val() : emailHeader;
        xmppHeader = $('#xmppHeader').val() != xmppHeader ? $('#xmppHeader').val() : xmppHeader;
        twitterHeader = $('#twitterHeader').val() != twitterHeader ? $('#twitterHeader').val() : twitterHeader;
        //update the headerMappings
        headerMappings["firstName"] = firstNameHeader;
        headerMappings["lastName"] = lastNameHeader;
        headerMappings["broadsoft"] = fixedLineHeader;
        headerMappings["sms"] = smsHeader;
        headerMappings["xmpp"] = xmppHeader;
        headerMappings["mail"] = emailHeader;
        headerMappings["twitter"] = twitterHeader;
    }

    //returns true or false based on the mediums selected in the checkbox: adapterTypes
    function shouldColumnBeIgnored(header) {
        for (var headerKey in headerMappings) {
            if (headerMappings[headerKey] == header
                && $.inArray(headerKey, getChannelsChecked()) != -1) {
                return false;
            }
        }
        return true;
    }

    //returns a string value of all the header query parameters added
    function appendHeaders() {
        return "&firstName=" + encodeURIComponent(firstNameHeader) +
            "&lastName=" + encodeURIComponent(lastNameHeader) +
            "&smsHeader=" + (($('#sms').is(":checked")) ? encodeURIComponent(smsHeader) : "") +
            "&callHeader=" + (($('#broadsoft').is(":checked")) ? encodeURIComponent(fixedLineHeader) : "") +
            "&emailHeader=" + (($('#mail').is(":checked")) ? encodeURIComponent(emailHeader) : "") +
            "&xmppHeader=" + (($('#xmpp').is(":checked")) ? encodeURIComponent(xmppHeader) : "") +
            "&twitterHeader=" + (($('#twitter').is(":checked")) ? encodeURIComponent(twitterHeader) : "");
    }

    //generates the report on the excel sheet for the reponse to the questions seen
    function genereateReport() {
        if (X_SESSIONID != null && X_SESSIONID != "") {
            app.showNotification("Generating report..", "");
            $.ajax({
                cache: false,
                crossDomain: true,
                contentType: 'application/json; charset=utf-8',
                beforeSend: function (request) {
                    request.setRequestHeader("X-SESSION_ID", X_SESSIONID);
                },
                url: './ASKFastRequestHandler.ashx/report',
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
            var data = new Array();
            var rowCounter = 0;
            var dataRowCounter = 1; //starts the data from the first row.
            data[0] = ["Timestamp", "Message Type", "Question", "Responder", "Medium", "Responder Name",
                "Status", "Response"];
            for (; rowCounter < response.length; rowCounter++) {
                var questionResponse = response[rowCounter];
                var rowData = new Array();
                //initialize with empty values
                for (var columnCounter = 0; columnCounter < data[0].length; columnCounter++) {
                    rowData[columnCounter] = "";
                }
                rowData[0] = getStringDateFromMilliseconds(questionResponse["timestamp"]);
                var questionMap = questionResponse["clipboardMap"];
                if (questionMap["question"] != null) {
                    var question = JSON.parse((questionMap["question"]));
                    var questionType = "";
                    switch (question["type"].toString().toLowerCase()) {
                        case 'comment':
                            questionType = 'Broadcast';
                            break;
                        case 'open':
                            questionType = 'Open';
                            break;
                        case 'closed':
                            questionType = 'Yes/No';
                            break;
                    }
                    rowData[1] = questionType;
                    var questionText = question["question_text"];
                    if (questionText != null) {
                        questionText = questionText.replace('text://', '');
                        questionText = decodeURIComponent(questionText);
                    }
                    rowData[2] = questionText;
                    //ignore this question if the text doesnt match
                    if ($('#matchQuestionCheckBox').is(":checked") && $('#message').val()
                        && $('#message').val() != decodeURIComponent(questionText)) {
                        continue;
                    }
                }
                rowData[3] = questionMap["responder"];
                rowData[4] = questionMap["adapterType"];
                rowData[5] = questionMap["responder_name"];
                rowData[6] = questionMap["status"];
                rowData[7] = questionMap["answer_text"];
                data[dataRowCounter++] = rowData;
            }
            if (data.length > 1) {
                Office.context.document.setSelectedDataAsync(data, { coercionType: "matrix" },
                    function (asyncResult) {
                        if (asyncResult.status === "failed") {
                            app.showNotification(asyncResult.error.name, asyncResult.error.message);
                        }
                    });
            }
            else {
                app.showNotification("Info", "Reports found but none matching the question: " + $('#message').val());
            }
        }
        else {
            app.showNotification("Info", "No reports found.")
        }
    }

    function getStringDateFromMilliseconds(milliseconds) {
        var date = new Date(milliseconds);
        return date.toDateString() + " " + date.toLocaleTimeString();
    }

    function toggleMatchingQuestionText() {
        if ($('#matchQuestionCheckBox').is(":checked")) {
            $('#matchingQuestionText').text("Fetchs reports for current message.");
        }
        else {
            $('#matchingQuestionText').text("Fetchs all reports.");
        }
    }

    //store the userName and password in the localstorage
    function supports_html5_storage() {
        try {
            return 'localStorage' in window && window['localStorage'] !== null;
        } catch (e) {
            return false;
        }
    }

    //store all the entered info in the local storage
    function storeDataInLocalStorage() {
        if (supports_html5_storage()) {
            localStorage.setItem("message", $('#message').val());
            localStorage.setItem("questionType", $('#questionType').val());
            localStorage.setItem("language", $('#language').val());
            localStorage.setItem("xmpp", $('#xmpp').is(":checked"));
            localStorage.setItem("mail", $('#mail').is(":checked"));
            localStorage.setItem("broadsoft", $('#broadsoft').is(":checked"));
            localStorage.setItem("sms", $('#sms').is(":checked"));
            localStorage.setItem("twitter", $('#twitter').is(":checked"));
            localStorage.setItem("senderId", $('#senderId').val());
            localStorage.setItem("subject", $('#subject').val());
            localStorage.setItem("firstNameHeader", $('#firstNameHeader').val());
            localStorage.setItem("lastNameHeader", $('#lastNameHeader').val());
            localStorage.setItem("xmppHeader", $('#xmppHeader').val());
            localStorage.setItem("smsHeader", $('#smsHeader').val());
            localStorage.setItem("fixedLineHeader", $('#fixedLineHeader').val());
            localStorage.setItem("emailHeader", $('#emailHeader').val());
            localStorage.setItem("twitterHeader", $('#twitterHeader').val());
        }
    }

    //load all saved info from the local storage
    function loadDataFromLocalStorage() {
        if (supports_html5_storage()) {
            $('#message').val(localStorage.getItem("message"));
            $('#questionType').val(localStorage.getItem("questionType"));
            $('#language').val(localStorage.getItem("language"));
            if (localStorage.getItem("xmpp")) {
                $("#xmpp").prop("checked", localStorage.getItem("xmpp") === 'true');
            }
            if (localStorage.getItem("mail")) {
                $("#mail").prop("checked", localStorage.getItem("mail") === 'true');
            }
            if (localStorage.getItem("broadsoft")) {
                $("#broadsoft").prop("checked", localStorage.getItem("broadsoft") === 'true');
            }
            if (localStorage.getItem("sms")) {
                $("#sms").prop("checked", localStorage.getItem("sms") === 'true');
            }
            if (localStorage.getItem("twitter")) {
                $("#twitter").prop("checked", localStorage.getItem("twitter") === 'true');
            }
            $('#senderId').val(localStorage.getItem("senderId"));
            $('#subject').val(localStorage.getItem("subject"));
            if (localStorage.getItem("firstNameHeader")) {
                $('#firstNameHeader').val(localStorage.getItem("firstNameHeader"));
            }
            if (localStorage.getItem("lastNameHeader")) {
                $('#lastNameHeader').val(localStorage.getItem("lastNameHeader"));
            }
            if (localStorage.getItem("xmppHeader")) {
                $('#xmppHeader').val(localStorage.getItem("xmppHeader"));
            }
            if (localStorage.getItem("smsHeader")) {
                $('#smsHeader').val(localStorage.getItem("smsHeader"));
            }
            if (localStorage.getItem("fixedLineHeader")) {
                $('#fixedLineHeader').val(localStorage.getItem("fixedLineHeader"));
            }
            if (localStorage.getItem("emailHeader")) {
                $('#emailHeader').val(localStorage.getItem("emailHeader"));
            }
            if (localStorage.getItem("twitterHeader")) {
                $('#twitterHeader').val(localStorage.getItem("twitterHeader"));
            }
        }
    }
})();