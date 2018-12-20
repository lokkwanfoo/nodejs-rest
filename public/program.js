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

    getOneDriveFiles("/api/templates");

        $("#test").click(function () {
            getOneDriveFiles("/clear");
        });

        $("#gets").click(function () {
            getOneDriveFiles("/api/templates");
        });

        $("#get").click(function () {
            getTemplateWithoutAuth("/api/template", "01KPZU6TPBTUP5KYM5XNF3VTZIVUKQKDXY")

        });

        $("#dialog").click(function () {
            Office.context.ui.displayDialogAsync('https://localhost:3000/profile.html', {height: 50, width: 50}, function(asyncResult) {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                Office.context.auth.getAccessTokenAsync({forceConsent: false},
                    function (result) {
                        localStorage.setItem("accessToken", result.value);
                    });
            });
        });


        $("#paste").click(function () {

            Word.run(function (context) {
                // const logo = context.document.contentControls.getByTag("logo").getFirst();

                // logo.insertInlinePictureFromBase64(image, "Replace");

                // const logo = context.document.contentControls.getByTag("bankaccount").getFirst();
                // logo.insertInlinePictureFromBase64(profile.image, "Start");

                const subject = context.document.contentControls.getByTag("subject").getFirst();
                subject.insertText("Rekening", "Replace");

                return context.sync();
            })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });

            
        });

        $("#clear").click(function () {
            Word.run(function (context) {

                context.document.body.clear();

                return context.sync();
            })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        });

    });
}
    var templates = [];
    var image;
    var profile;
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function processMessage(arg) {
        profile = JSON.parse(arg.message)
        getTemplateWithoutAuth("/api/template", "01KPZU6TPBTUP5KYM5XNF3VTZIVUKQKDXY");
    }

    function getBase64(file) {
        return new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.readAsDataURL(file);
          reader.onload = () => resolve(reader.result);
          reader.onerror = error => reject(error);
        });
    }

    function uploadPhoto(e) 
    {
        Word.run(function (context) {

            let photo = e.files[0] 

            getBase64(photo).then(
                data => 
                // context.document.body.insertInlinePictureFromBase64(data.substr(data.indexOf(',') + 1), "Start")
                image = data.substr(data.indexOf(',') + 1)
            );

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        
    }
    
    function getOneDriveFiles(apiURLsegment, nameDocument) {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithoutAuthChallenge(apiURLsegment, nameDocument);
    }   

    function postOneDriveFiles(apiURLsegment, nameDocument) {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        postDataWithoutAuthChallenge(apiURLsegment, nameDocument);
    }  

    function generateTemplate(profile, template) {

        Word.run(function (context) {

            // const name = context.document.contentControls.getByTag("name").getFirst();
            // name.insertText(profile.name, "Replace");

            var wordDocument = context.application.createDocument(template);

            var contentControls = context.document.contentControls;

    // Queue a command to load the content controls collection.
    console.log(contentControls.load('name'))


            if (wordDocument) {
                
                // const name = context.document.contentControls.getByTag("name").getFirst();
                // name.insertText(profile.name, "Replace");
                wordDocument.open();
                
            }

            return context.sync().then(function() {
                if (wordDocument.contentControls.items.length !== 0) {
                    for (var i = 0; i < wordDocument.contentControls.items.length; i++) {
                        console.log(wordDocument.contentControls.items[i].id);
                        console.log(wordDocument.contentControls.items[i].text);
                        console.log(wordDocument.contentControls.items[i].tag);
                    }
                } else {
                    console.log('No content controls in this document.');
                }
            });
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }

    function test(test, a) {
        Word.run(function (context) {

            var wordDocument = context.application.createDocument(a);

            if (wordDocument) {
                wordDocument.open();
             }

            // context.document.body.insertFileFromBase64(a, "Start");

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    // Called in the first attempt to use the on-behalf-of flow. The assumption
    // is that single factor authentication is all that is needed.
    async function getDataWithoutAuthChallenge(apiURLsegment, nameDocument) {
        Office.context.auth.getAccessTokenAsync({forceConsent: false},
            function (result) {
                if (result.status === "succeeded") {
                    accessToken = result.value;
                    getData(apiURLsegment, accessToken, nameDocument);
                }
                else {
                    // test(result.error.message)
                    console.log(result)
                    handleClientSideErrors(result);
                }
            });
    }

    async function postDataWithoutAuthChallenge(apiURLsegment, nameDocument) {
        Office.context.auth.getAccessTokenAsync({forceConsent: false},
            function (result) {
                if (result.status === "succeeded") {
                    accessToken = result.value;
                    postData(apiURLsegment, accessToken, nameDocument);
                }
                else {
                    test(result.error.message)
                    console.log(result)
                    handleClientSideErrors(result);
                }
            });
    }
    
    // Calls the specified URL or route (in the same domain as the add-in)
    // and includes the specified access token.
    async function getTemplateWithoutAuth(apiURLsegment, path) {

        Office.context.auth.getAccessTokenAsync({forceConsent: false},
            function (result) {
                if (result.status === "succeeded") {
                    accessToken = result.value;
                    getTemplate(apiURLsegment, accessToken, path);
                }
                else {
                    // test(result.error.message)
                    console.log(result)
                    handleClientSideErrors(result);
                }
            });
    }

    async function getTemplate(relativeUrl, accessToken, path) {

        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken, "Path": path},
            type: "GET",
            // Turn off caching when debugging to force a fetch of data
            // with each call.
            cache: false
        })
        .done(function (result) {
            generateTemplate(profile, result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
            test("error")
            console.log(result.responseJSON.error);
        });
    }

    // Calls the specified URL or route (in the same domain as the add-in)
    // and includes the specified access token.
    async function getData(relativeUrl, accessToken, path) {

        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken, "Path": path},
            type: "GET",
            // Turn off caching when debugging to force a fetch of data
            // with each call.
            cache: false
        })
        .done(function (result) {
            if (Array.isArray(result)){ 
                templates = result;
                console.log(result)
            } 
            // If the result contains 'capolids', then it is the Claims string,
            // not the data.
            else if (result[0].indexOf('capolids') !== -1) {
                result[0] = JSON.parse(result[0])
                getDataUsingAuthChallenge(result[0]);
            } 
            // else if (typeof result === 'string') {
            //     test("", result);
            // }
            else {
                console.log(result)
                test("", result);
                templates = result;
            }
        })
        .fail(function (result) {
            handleServerSideErrors(result);
            test("error")
            console.log(result.responseJSON.error);
        });
    }

    function postData(relativeUrl, accessToken, path) {

        var profile = {
            "name": "T. Tester",
            "initials": "t.t.",
            "phonenumber":"0201234567",
            "faxnumber": "",
            "mobilenumber":"0687654321",
            "email": "t.tester@vandoorne.nl",
            "roleDutch": "Advocaat",
            "roleEnglish": "Attorney",
            "roleGerman": "Rechtsanwalt",
            "image": image
        }

        var profileName = 'Persoonsprofiel';

        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken, "Path": path, "profileName": profileName},
            type: "POST",
            // Turn off caching when debugging to force a fetch of data
            // with each call.
            cache: false,
            data: profile
        })
        .done(function (result) {
            console.log(result)
            // test("success", result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
            test("error")
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
                showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.' +
                'Other kinds of accounts, like corporate domain accounts do not work.']);
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
