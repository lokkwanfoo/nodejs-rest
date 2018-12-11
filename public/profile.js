Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
        $("#getGraphAccessTokenButton").click(function () {
            Office.context.ui.messageParent('{"a":2}');
        });

        $("#getProfiles").click(function () {
            getData("/api/profiles", accessToken);
        });

        $("#getProfile").click(function () {
            getData("/api/profile", accessToken, profiles[0].id);
        });

    });
}

var accessToken = localStorage.getItem("accessToken");
var profiles;
var profile;

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
            console.log(profiles)

            var str = '<ul>'

            profiles.forEach(function(profile) {
            str += '<li>'+ profile.name + '</li>';
            }); 

            str += '</ul>';
            document.getElementById("profiles").innerHTML = str;

        } 
        else {
            profile = result;
            console.log(result)
        }
    })
    .fail(function (result) {
        console.log(result.responseJSON.error);
    });
}

function postData(relativeUrl, accessToken, path) {

    var profile = {
        "name": "asdasdsd",
        "initials": "asdasdasd",
        "phonenumber":"",
        "faxnumber": "",
        "mobilenumber":"",
        "email": "",
        "roleDutch": "",
        "roleEnglish": "",
        "roleGerman": "",
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
    })
    .fail(function (result) {
        console.log(result.responseJSON.error);
    });
}

