Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
        // getData("/api/profiles", accessToken);

        if (JSON.parse(localStorage.getItem("profiles").length != 0)) {
            profiles = JSON.parse(localStorage.getItem("profiles"));
            console.log(profiles)
            fillList(profiles);
        }
        
        $("#close").click(function () {
            console.log(profile)
            Office.context.ui.messageParent(JSON.stringify(profile));
        });

        $("#getProfiles").click(function () {
            getData("/api/profiles", accessToken);
        });

        $("#profiles").click(function () {
            console.log(document.getElementById("profiles").id)
            console.log(profiles[document.getElementById("profiles").value].id)
            // fillForm();
            // getData("/api/profile", accessToken, document.getElementById("profiles").value);
        });


        $("#saveProfile").click(function () {
            postData("/api/profile", accessToken);
        });

        $("#deleteProfile").click(function () {
            getData("/delete/profile", accessToken, profiles[document.getElementById("profiles").value].id).then(function() {
                getData("/api/profiles", accessToken).then(function(result) {
                    profile = result;
                    console.log(result)
                    fillList(result);
                });
            });
        });

    });
}
// 
var accessToken = localStorage.getItem("accessToken");
var image;
var profiles;
var profile;
// 
function fillList(array) {
    $("#profiles").empty();
    var select = document.getElementById("profiles"); 
    if (!!array) {
        for (var i in array) {
            var el = document.createElement("option");
            el.textContent = array[i].profileName;
            el.value = array[i].id;
            select.appendChild(el);
        }
    } 
}

function fillForm(profile) {
    document.getElementById("name").value = profile.name ? profile.name : '';
    document.getElementById("initials").value = profile.initials ? profile.initials : '';
    document.getElementById("phonenumber").value = profile.phonenumber ? profile.phonenumber : '';
    document.getElementById("faxnumber").value = profile.faxnumber ? profile.faxnumber : '';
    document.getElementById("mobilenumber").value = profile.mobilenumber ? profile.mobilenumber : '';
    document.getElementById("emailaddress").value = profile.emailaddress ? profile.emailaddress : '';
    document.getElementById("roleDutch").value = profile.roleDutch ? profile.roleDutch : '';
    document.getElementById("roleEnglish").value = profile.roleEnglish ? profile.roleEnglish : '';
    document.getElementById("roleGerman").value = profile.roleGerman ? profile.roleGerman : '';
    document.getElementById("profileName").value = profile.profileName ? profile.profileName : '';
}

function readForm() {
    return new Promise(function(resolve, reject) {
        profile = {
            "name":  document.getElementById("name").value ? document.getElementById("name").value : '',
            "initials": document.getElementById("initials").value ? document.getElementById("initials").value : '',
            "phonenumber": document.getElementById("phonenumber").value ? document.getElementById("phonenumber").value : '',
            "faxnumber": document.getElementById("faxnumber").value ? document.getElementById("faxnumber").value : '',
            "mobilenumber": document.getElementById("mobilenumber").value ? document.getElementById("mobilenumber").value : '',
            "emailaddress": document.getElementById("emailaddress").value ? document.getElementById("emailaddress").value : '',
            "roleDutch": document.getElementById("roleDutch").value ? document.getElementById("roleDutch").value : '',
            "roleEnglish": document.getElementById("roleEnglish").value ? document.getElementById("roleEnglish").value : '',
            "roleGerman": document.getElementById("roleGerman").value ? document.getElementById("roleGerman").value : '',
            "standard": false,
            "profileName": document.getElementById("profileName").value
        }
        resolve(profile);
    })
}

function getData(relativeUrl, accessToken, path) {
    return new Promise(function(resolve, reject) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken, "Path": path},
            type: "GET",
            // Turn off caching when debugging to force a fetch of data
            // with each call.
            cache: false
        })
        .done(function (result) {
            resolve(result);
        })
        .fail(function (result) {
            reject(result)
            console.log(result.responseJSON.error);
        });
    })  
}

function postData(relativeUrl, accessToken, path) {

    if (!!document.getElementById("profileName").value)  {

        readForm().then(function(result) {
            $.ajax({
                url: relativeUrl,
                headers: { "Authorization": "Bearer " + accessToken, "Path": path, "profilename": document.getElementById("profileName").value},
                type: "POST",
                // Turn off caching when debugging to force a fetch of data
                // with each call.
                cache: false,
                data: result
            })
            .done(function (result) {
                getData("/api/profiles", accessToken).then(function(result) {
                    profiles = result;
                    fillList(result);
                });
            })
            .fail(function (result) {
                console.log(result.responseJSON.error);
            });
    
        })

    }
    
}

