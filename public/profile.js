Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
        // getData("/api/profiles", accessToken);

        if (JSON.parse(localStorage.getItem("profiles").length != 0)) {
            profiles = JSON.parse(localStorage.getItem("profiles"));
            fillList(profiles);
            getData("/api/locations", accessToken).then(function(result) {
                fillDropdown(result);
                locations = result;
            })
        }
        
        $("#close").click(function () {
            Office.context.ui.messageParent("");
        });

        $("#profiles").click(function () {
            fillForm(profiles[document.getElementById("profiles").value]);
        });

        $("#location").click(function () {
            console.log(document.getElementById("location").value)
        });

        $("#saveProfile").click(function () {
            addProfile();
        });

        $("#getLocations").click(function () {
            getData("/api/locations", accessToken).then(function(result) {
                fillDropdown(result);
            })
        });

        $("#deleteProfile").click(function () {
            for (var i in profiles) {
                if (i == document.getElementById("profiles").value) {
                    profiles.splice(i, 1);
                    postData("/api/profile", accessToken);
                }
            }
        });

        $("#makeDefault").click(function () {
            makeDefault();
        });

    });
}

var accessToken = localStorage.getItem("accessToken");
var profiles = [];
var profile;
var locations;

function fillList(array) {
    $("#profiles").empty();
    var select = document.getElementById("profiles"); 
    if (!!array) {
        for (var i in array) {
            var el = document.createElement("option");
            el.textContent = array[i].profileName;
            el.value = i;
            select.appendChild(el);
        }
    } 
}

function fillDropdown(array) {
    $("#location").empty();
    var select = document.getElementById("location"); 
    if (!!array) {
        for (var i in array) {
            var el = document.createElement("option");
            el.textContent = array[i].title;
            el.value = array[i].id;
            select.appendChild(el);
        }
    } 
}

function fillForm(profile) {
    var profileProps = Object.getOwnPropertyNames(profile);
    for (var i in profileProps) {
        if (!!document.getElementById(profileProps[i])) {
            document.getElementById(profileProps[i]).value = profile[profileProps[i]] ? profile[profileProps[i]] : '';
        }
    }
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
            "location": document.getElementById("location").value ? document.getElementById("location").value : '',
            "roleDutch": document.getElementById("roleDutch").value ? document.getElementById("roleDutch").value : '',
            "roleEnglish": document.getElementById("roleEnglish").value ? document.getElementById("roleEnglish").value : '',
            "roleGerman": document.getElementById("roleGerman").value ? document.getElementById("roleGerman").value : '',
            "default": document.getElementById("default").value ? document.getElementById("default").value : '',
            "profileName": document.getElementById("profileName").value
        }
        resolve(profile);
    })
}

function addProfile() {
    var exists = false;
    if (!!document.getElementById("profileName").value)  {
        readForm().then(function(result) {
            if (profiles.length != 0) {
                for (var i = 0; i < profiles.length; i++) {
                    if (profiles[i].profileName === result.profileName) {
                        profiles[i] = result;
                        exists = true;
                    } 
                    if (i == profiles.length -1 && !exists) {
                        profiles.push(result);
                    }
                }
                postData("/api/profile", accessToken)
            } else {
                profiles.push(result);
                postData("/api/profile", accessToken)
            }
        })
    }
}

function makeDefault() {
    for (var i in profiles) {
        profiles[i].default = false;
    }
    profiles[document.getElementById("profiles").value].default = true;
    postData("/api/profile", accessToken)
}

function getData(relativeUrl, accessToken) {
    return new Promise(function(resolve, reject) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken},
            type: "GET",
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

function postData(relativeUrl, accessToken) {
    $.ajax({
        url: relativeUrl,
        headers: { "Authorization": "Bearer " + accessToken},
        type: "POST",
        cache: false,
        data: {value: profiles}
    })
    .done(function (result) {
        getData("/api/profiles", accessToken).then(function(result) {
            profiles = JSON.parse(result);
            fillList(profiles);
        });
    })
    .fail(function (result) {
        console.log(result.responseJSON.error);
    });
}

