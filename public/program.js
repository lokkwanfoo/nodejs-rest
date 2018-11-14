// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/*
    This file provides functions to get ask the Office host to get an access token to the add-in
	and to pass that token to the server to get Microsoft Graph data.
*/

Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
        $("#getGraphAccessTokenButton").click(function () {
                getOneDriveFiles();
        });
    });
}

    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithoutAuthChallenge();
    }   

    // Called in the first attempt to use the on-behalf-of flow. The assumption
    // is that single factor authentication is all that is needed.
    function getDataWithoutAuthChallenge() {
        Office.context.auth.getAccessTokenAsync({forceConsent: false},
            function (result) {
                if (result.status === "succeeded") {
                    accessToken = result.value;
                    getData("/api/me", accessToken, "/displayName");
                }
                else {
                    handleClientSideErrors(result);
                    console.log("Code: " + result.error.code);
                    console.log("Message: " + result.error.message);
                    console.log("name: " + result.error.name);
                    document.getElementById("getGraphAccessTokenButton").disabled = true;
                }
            });
    }

    // Calls the specified URL or route (in the same domain as the add-in)
    // and includes the specified access token.
    function getData(relativeUrl, accessToken, path) {

        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken},
            path: path,
            type: "GET",
            // Turn off caching when debugging to force a fetch of data
            // with each call.
            cache: false
        })
        .done(function (result) {
            /*
              If the Microsoft Graph target requests addtional authentication
              factor(s), the result will not be data. It will be a Claims
              JSON telling AAD what addtional factors the user must provide.
              Start a new sign-on that passes this Claims string to AAD so that
              it will provide the needed prompts.
            */

            // If the result contains 'capolids', then it is the Claims string,
            // not the data.
            console.log(result)
            if (result[0].indexOf('capolids') !== -1) {
                result[0] = JSON.parse(result[0])
                getDataUsingAuthChallenge(result[0]);
            } else {
                showResult(result);
            }
        })
        .fail(function (result) {
            handleServerSideErrors(result);
            console.log(result.responseJSON.error);
        });
    }

    // Called to trigger a second sign-on in which the user will be prompted
    // to provide additional authentication factor(s). The authChallengeString
    // parameter tells AAD what factor(s) it should prompt for.
    function getDataUsingAuthChallenge(authChallengeString) {
        Office.context.auth.getAccessTokenAsync({authChallenge: authChallengeString},
            function (result) {
                if (result.status === "succeeded") {
                    accessToken = result.value;
                    getData("/api/onedriveitems", accessToken);
                }
                else {
                    console.log("Code: " + result.error.code);
                    console.log("Message: " + result.error.message);
                    console.log("name: " + result.error.name);
                    document.getElementById("getGraphAccessTokenButton").disabled = true;
                }
            });
    }

    // Displays the data, assumed to be an array.
    function showResult(data) {
        for (var i = 0; i < data.length; i++) {
            $('#file-list').append('<li class="ms-ListItem">' +
            '<span class="ms-ListItem-secondaryText">' +
            '<span class="ms-fontColor-themePrimary">' + data[i] + '</span>' +
            '</span></li>');
        }
    }

    function handleClientSideErrors(result) {

        switch (result.error.code) {
    
            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor. 

            case 13001:
                getDataWithToken({ forceAddAccount: true });
                break;
    
            // TODO3: Handle the case where the user's sign-in or consent was aborted.

            case 13002:
                if (timesGetOneDriveFilesHasRun < 2) {
                    showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
                } else {
                    logError(result);
                }          
                break; 
    
            // TODO4: Handle the case where the user is logged in with an account that is neither work or school, 
            //        nor Micrososoft Account.

            case 13003: 
                showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
                break;   
    
            // TODO5: Handle an unspecified error from the Office host.

            case 13006:
                showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
                break;  
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.

            case 13007:
                showResult(['That operation cannot be done at this time. Please try again later.']);
                break; 
    
            // TODO7: Handle the case where the user tiggered an operation that calls `getAccessTokenAsync` 
            //        before a previous call of it completed.

            case 13008:
                showResult(['Please try that operation again after the current operation has finished.']);
                break;
    
            // TODO8: Handle the case where the add-in does not support forcing consent.

            case 13009:
                if (triedWithoutForceConsent) {
                    showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
                } else {
                    getDataWithToken({ forceConsent: false });
                }
                break;
    
            // TODO9: Log all other client errors.

            default:
                logError(result);
                break;
        }
    }

    function handleServerSideErrors(result) {

        // TODO10: Handle the case where AAD asks for an additional form of authentication.

        if (result.responseJSON.error.innerError
                && result.responseJSON.error.innerError.error_codes
                && result.responseJSON.error.innerError.error_codes[0] === 50076){
            getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
        }
    
        // TODO11: Handle the case where consent has not been granted, or has been revoked.

        else if (result.responseJSON.error.innerError
                && result.responseJSON.error.innerError.error_codes
                && result.responseJSON.error.innerError.error_codes[0] === 65001){
            showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
            /*
                THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
                OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
                THE FOLLOWING LINE.
            */
            // getDataWithToken({ forceConsent: true });
        }
    
        // TODO12: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        else if (result.responseJSON.error.innerError
                && result.responseJSON.error.innerError.error_codes
                && result.responseJSON.error.innerError.error_codes[0] === 70011){
            showResult(['The add-in is asking for a type of permission that is not recognized.']);
        }
    
        // TODO13: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        else if (result.responseJSON.error.name
                && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
            showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
        }
    
        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        else if (result.responseJSON.error.name
                && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
            if (!timesMSGraphErrorReceived) {
                timesMSGraphErrorReceived = true;
                timesGetOneDriveFilesHasRun = 0;
                triedWithoutForceConsent = false;
                getOneDriveFiles();
            } else {
                logError(result);
            }        
        }
    
        // TODO15: Log all other server errors.

        else {
            logError(result);
        }
    }

    function logError(result) {
        console.log(result);
    }
