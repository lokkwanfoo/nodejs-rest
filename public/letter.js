Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
        getData("/api/profiles", accessToken);

        $("#close").click(function () {
            console.log(profile)
            Office.context.ui.messageParent(JSON.stringify(profile));
            window.close();
        });

        $("#getProfiles").click(function () {
            getData("/api/profiles", accessToken);
            var anyncFunction = async(function (callback) {
                console.log(getData("/api/profiles", accessToken))
                callback();
            });
            
        });

        $("#profiles").click(function () {
            getData("/api/profile", accessToken, document.getElementById("profiles").value);
        });


        $("#saveProfile").click(function () {
            postData("/api/profile", accessToken);
        });

        $("#deleteProfile").click(function () {
            getData("/delete/profile", accessToken, document.getElementById("profiles").value);
        });

    });
}

var accessToken = localStorage.getItem("accessToken");
var image;
var profiles;
var profile;

function async(your_function, callback) {
    setTimeout(function() {
        your_function();
        if (callback) {callback();}
    }, 0);
}

function getData(relativeUrl, accessToken, path) {

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
            profiles = result;
            $("#profiles").empty();
            var select = document.getElementById("profiles"); 
            for (var i in profiles) {
                var el = document.createElement("option");
                el.textContent = profiles[i].name.substr(0, (profiles[i].name.indexOf('.')));
                el.value = profiles[i].id;
                select.appendChild(el);
            }
        } 
        else {
            profile = JSON.parse(result);
            document.getElementById("name").value = profile.name;
            document.getElementById("initials").value = profile.initials;
            document.getElementById("phonenumber").value = profile.phonenumber;
            document.getElementById("faxnumber").value = profile.faxnumber;
            document.getElementById("mobilenumber").value = profile.mobilenumber;
            document.getElementById("emailaddress").value = profile.emailaddress;
            document.getElementById("location").value = profile.location;
            document.getElementById("profilename").value = $("#profiles option:selected").text();
            getData("/api/profiles", accessToken);
            return result;
        }
    })
    .fail(function (result) {
        console.log(result.responseJSON.error);
    });
}

function postData(relativeUrl, accessToken, path) {

    var profile = {
        "name":  document.getElementById("name").value,
        "initials": document.getElementById("initials").value,
        "phonenumber": document.getElementById("phonenumber").value,
        "faxnumber": document.getElementById("faxnumber").value,
        "mobilenumber": document.getElementById("mobilenumber").value,
        "emailaddress": document.getElementById("emailaddress").value,
        "roleDutch": "Advocaat",
        "roleEnglish": "Attorney",
        "roleGerman": "Rechtsanwalt"
    }

    var profileName;

    $.ajax({
        url: relativeUrl,
        headers: { "Authorization": "Bearer " + accessToken, "Path": path, "profileName": document.getElementById("profilename").value},
        type: "POST",
        // Turn off caching when debugging to force a fetch of data
        // with each call.
        cache: false,
        data: profile
    })
    .done(function (result) {
        getData("/api/profiles", accessToken);
        console.log(result)
    })
    .fail(function (result) {
        console.log(result.responseJSON.error);
    });
}

