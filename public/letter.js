Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.

    getData("/api/languagepack", accessToken).then(function(result) {
        console.log(JSON.parse(result))
        languagepack = JSON.parse(result)
        fillDropdown(languagepack, "language");
        var languageProps = Object.keys(languagepack[0]);
        fillOtherDropdown(languagepack[0][languageProps[i]], languageProps)
    })

        $("#close").click(function () {
            postData();
            Office.context.ui.messageParent(JSON.stringify(letter));
            window.close();
        });

        $("#cancel").click(function () {
            dialog.close();
        });

        $("#language").click(function () {
            var languageProps = Object.keys(languagepack[document.getElementById("language").value]);
            // console.log(languageProps)
            // console.log(languagepack[document.getElementById("language").value])
            for (var i in languageProps) {
                if (languageProps[i] != "language") {
                    // fillOtherDropdown(languagepack[document.getElementById("language").value], languageProps[i])
                }
            }
        });

    });
}

var accessToken = localStorage.getItem("accessToken");
var letter;
var languagepack;

function fillDropdown(array, value) {
    console.log(array)
    $("#" + value).empty();
    var select = document.getElementById(value); 
    if (!!array) {
        for (var i in array) {
            var el = document.createElement("option");
            el.textContent = array[i][value];
            el.value = i;
            select.appendChild(el);
        }
    } 
}

function fillOtherDropdown(array, props) {
    // console.log(array)
    // console.log(props)
    var select = document.getElementById(props)
    for (var i in array) {
        for (var j in array[i]) {
            var el = document.createElement("option");
        el.textContent = array[i];
        el.value = i;
        select.appendChild(el)
        }
        
    }

    // for (var i in array) {
    //     $("#" + props[i]).empty();
    //     var select = document.getElementById(props[i]); 
    //     if (!!array) {
    //         var el = document.createElement("option");
    //         el.textContent = array[i];
    //         el.value = i;
    //         select.appendChild(el);
    //     }     
    // }   
}

function getData(relativeUrl, accessToken) {
    return new Promise(function(resolve, reject) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken},
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

function postData() {

    letter = {
        "nameaddress":  document.getElementById("nameaddress").value,
        "yourReference": document.getElementById("yourReference").value,
        "ourReference": document.getElementById("ourReference").value,
        "subject": document.getElementById("subject").value
        // "header": document.getElementById("header").value,
        // "closer": document.getElementById("closer").value,
        // "signer": document.getElementById("signer").value
    }
    
}

