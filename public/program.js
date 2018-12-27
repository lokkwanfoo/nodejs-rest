Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.

    getOneDriveFiles("/api/templates").then(function(result) {
        templates = result;
    })

    getToken();
    
        $("#test").click(function () {
            getOneDriveFiles("/api/profiles").then(function(result) {
                console.log(result)
            })
        });

        $("#gets").click(function () {
            console.log("asd")
            getOneDriveFiles("/api/templates").then(function(result) {
                templates = result;
            });
        });

        $("#get").click(function () {
            clearContentControls();
            getOneDriveFiles("api/template", templates[0].url + "/" + templates[0].name).then(function(result){
                template = result;
                generateTemplate(letter, template, profile);
            });
        });

        $("#dialog").click(function () {
            openDialog('profile.html', 85, 50);
        });

        $("#letter").click(function () {
            openDialog('letter.html', 85, 75);
        });

        $("#paste").click(function () {
            generateTemplateNewFile(profile, template);    
        });

        $("#clear").click(function () {
            clearContentControls();
        });
    })
};

    var templates = [];
    var template;
    var profile;
    var profiles;
    var image;

    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    var profileStructure = {
        "emailaddress": "",
        "faxnumber": "",
        "initials": "",
        "mobilenumber": "",
        "name": "",
        "phonenumber": "",
        "roleDutch": "",
        "roleEnglish": "",
        "roleGerman": ""
    }

    var letterStructure = {
        "nameaddress": "",
        "yourReference": "",
        "ourReference": "",
        "subject": "",
        "header": "",
        "closer": "",
        "signer": ""
    }

    function getToken() {
        return new Promise(function(resolve, reject) {
            Office.context.auth.getAccessTokenAsync({forceConsent: false},
                function (result) {
                    if (result.status === "succeeded") {
                        localStorage.setItem("accessToken", result.value);
                    }
                    else {
                        reject(result);
                        console.log(result)
                        handleClientSideErrors(result);
                    }
                });
        })
    }

    function openDialog(url, height, width) {
        if (url == "profile.html") {
            getOneDriveFiles("/api/profiles").then(function(result) {
                localStorage.setItem("profiles", JSON.stringify(result));
                Office.context.ui.displayDialogAsync("https://localhost:3000/" + url, {height: height, width: width, displayInIframe: true, promptBeforeOpen: false}, function(asyncResult) {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                });
            });
        }
    }

    function processMessage(arg) {
        dialog.close();
        if (isEquivalent(letterStructure, JSON.parse(arg.message))) {
            
            letter = JSON.parse(arg.message)
            clearContentControls();
            getOneDriveFiles("api/template", templates[0].url + "/" + templates[0].name).then(function(result){
                template = result;
                console.log(result)
                generateTemplate(letter, template, profile);
            });

        }
        if (isEquivalent(profileStructure, JSON.parse(arg.message))) {
            profile = JSON.parse(arg.message)
        }
    }

    function isEquivalent(a, b) {
        // Create arrays of property names
        var aProps = Object.getOwnPropertyNames(a);
        var bProps = Object.getOwnPropertyNames(b);
    
        // If number of properties is different,
        // objects are not equivalent
        if (aProps.length != bProps.length) {
            return false;
        }

        for (i in a) {
            if (!b.hasOwnProperty(i)) {
                return false;
            }
        }
    
        // If we made it this far, objects
        // are considered equivalent
        return true;
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

    function clearContentControls() {

        Word.run(function (context) {

            var contentControls = context.document.contentControls;

            contentControls.load();

            return context.sync().then(function () {

                if (contentControls.items.length === 0) {
                    console.log("There isn't a content control in this document.");
                    context.document.body.clear();
                } else {
                    for (var i in contentControls.items) {
                    // Queue a command to clear the contents of the first content control.
                        contentControls.items[i].delete(false);
                    }
                    
                    context.document.body.clear();
                    // Synchronize the document state by executing the queued commands, 
                    // and return a promise to indicate task completion.
                    return context.sync().then(function () {
                        console.log('Content control cleared of contents.');
                    });      
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
    
    function getOneDriveFiles(apiURLsegment, nameDocument) {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        return new Promise(function(resolve) {
            getDataWithoutAuthChallenge(apiURLsegment, nameDocument)
            .then(function(result) {
                    resolve(result);
            });
        })
    } 

    function postOneDriveFiles(apiURLsegment, nameDocument) {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        postDataWithoutAuthChallenge(apiURLsegment, nameDocument);
    }  

    function generateTemplateNewFile(profile, template) {

        Word.run(function (context) {
    
            var wordDocument = context.application.createDocument(template);
            
            if (wordDocument) {
                
                // const name = context.document.contentControls.getByTag("name").getFirst();
                // name.insertText(profile.name, "Replace");
                wordDocument.open();
                contentControls = wordDocument.contentControls;
            }
            return context.sync().then(function() {
                for (var i in contentControls) {
                    console.log(contentControls[i])
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

    function generateTemplate(letter, template, profile) {

        Word.run(function (context) {
            context.document.body.insertFileFromBase64(template, "Start")
        
            return context.sync().then(function() {
                var tempNameSpace = {};
                for (var i in letter) {
                    if (letter[i] ) {
                        tempNameSpace[i] = context.document.contentControls.getByTag(i).getFirstOrNullObject();
                        tempNameSpace[i].insertText(letter[i], "Replace");
                    }
                }
                for (var i in profile) {
                    if (profile[i] ) {
                        tempNameSpace[i] = context.document.contentControls.getByTag(i).getFirstOrNullObject();
                        tempNameSpace[i].insertText(profile[i], "Replace");
                    }
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

    // Called in the first attempt to use the on-behalf-of flow. The assumption
    // is that single factor authentication is all that is needed.
    function getDataWithoutAuthChallenge(apiURLsegment, nameDocument) {
        return new Promise(function(resolve, reject) {
            Office.context.auth.getAccessTokenAsync({forceConsent: false},
                function (result) {
                    if (result.status === "succeeded") {
                        accessToken = result.value;
                        getDataWithPromise(apiURLsegment, accessToken, nameDocument)
                        .then(function(result) {
                                resolve(result)
                        })
                    }
                    else {
                        // test(result.error.message)
                        reject(result);
                        console.log(result)
                        handleClientSideErrors(result);
                    }
                });
        })
        
    }

    async function postDataWithoutAuthChallenge(apiURLsegment, nameDocument) {
        Office.context.auth.getAccessTokenAsync({forceConsent: false},
            function (result) {
                if (result.status === "succeeded") {
                    accessToken = result.value;
                    postData(apiURLsegment, accessToken, nameDocument);
                }
                else {
                    console.log(result)
                    handleClientSideErrors(result);
                }
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
            resolve(result)
            // if (relativeUrl == "/api/templates"){ 
            //     templates = result;
            // } 
            // else if (relativeUrl == "/api/profiles") {
            //     profiles = result;
            // }
            // else if (relativeUrl == "/api/locations") {
            //     console.log(result)
            // }
            // else {
            //     template = result;
            // }
        })
        .fail(function (result) {
            handleServerSideErrors(result);
            console.log(result.responseJSON.error);
        });
    }

    function getDataWithPromise(relativeUrl, accessToken, path) {
        return new Promise(function(resolve, reject) {

            $.ajax({
                url: relativeUrl,
                headers: { "Authorization": "Bearer " + accessToken, "Path": path},
                type: "GET",
                cache: false
            })
            .done(function (result) {
                resolve(result)
            })
            .fail(function (result) {
                reject(Error(result))
                handleServerSideErrors(result);
                console.log(result.responseJSON.error);
            });

        })
        
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
            // test("success", result);
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

    function showResult(data) {
        for (var i = 0; i < data.length; i++) {
            $('#file-list').append('<li class="ms-ListItem">' +
            '<span class="ms-ListItem-secondaryText">' +
              '<span class="ms-fontColor-themePrimary">' + data[i] + '</span>' +
            '</span></li>');
        }
    }

    function logError(result) {
        console.log(result);
    }

