Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.

        getData("/api/languagepack", accessToken).then(function(result) {
            profiles = JSON.parse(localStorage.getItem("profiles"));
            for (var i in profiles) {
                if (profiles[i].default == "true") {
                    profiles = swapElement(profiles, i, 0);
                    fillDropdown(profiles, "name")
                }
            }
            languagepack = JSON.parse(result)
            fillDropdown(languagepack, "language");
            var languageProps = Object.keys(languagepack[0]);
            fillOtherDropdown(languagepack[0], languageProps);
        })

        $("#save").click(function () {
            postData().then(function(result){
                Office.context.ui.messageParent(JSON.stringify(result));
            });
        });

        $("#cancel").click(function () {
            Office.context.ui.messageParent("");
        });

        $("#header").change(function () {
            var e = document.getElementById("header")
            console.log(document.getElementById("header").options[document.getElementById("header").selectedIndex].text)
        });

        $("#language").change(function () {
            var languageProps = Object.keys(languagepack[document.getElementById("language").value]);
            fillOtherDropdown(languagepack[document.getElementById("language").value], languageProps);
        });

    });
}

var accessToken = localStorage.getItem("accessToken");
var languagepack;
var profiles;

function swapElement(array, indexA, indexB) {
    var tmp = array[indexA];
    array[indexA] = array[indexB];
    array[indexB] = tmp;
    return array;
  }

function fillDropdown(array, value) {
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
    for (var i in props) {
        var select = document.getElementById(props[i])
        if (props[i] != "language") {
            $("#" + props[i]).empty();
            for (var j in array[props[i]]) {
                var el = document.createElement("option");
                el.textContent = array[props[i]][j];
                el.value = j;
                select.appendChild(el)
            }
        }
    } 
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
        });
    })  
}

function postData() {
    return new Promise(function(resolve, reject) {
        var letter = {
            "nameaddress":  document.getElementById("nameaddress").value ? document.getElementById("nameaddress").value : '',
            "yourReference": document.getElementById("yourReference").value ? document.getElementById("nameaddress").value : '',
            "yourReference": document.getElementById("yourReference").value ? document.getElementById("nameaddress").value : '',
            "ourReference": document.getElementById("ourReference").value ? document.getElementById("ourReference").value : '',
            "subject": document.getElementById("subject").value ? document.getElementById("subject").value : '',
            "header": document.getElementById("header").options[document.getElementById("header").selectedIndex].text ? document.getElementById("header").options[document.getElementById("header").selectedIndex].text : '',
            "footer": document.getElementById("footer").options[document.getElementById("footer").selectedIndex].text ? document.getElementById("footer").options[document.getElementById("footer").selectedIndex].text : '',
            "status": document.getElementById("status").options[document.getElementById("status").selectedIndex].text ? document.getElementById("status").options[document.getElementById("status").selectedIndex].text : '',
            "sendOption" : document.getElementById("sendOption").options[document.getElementById("sendOption").selectedIndex].text ? document.getElementById("sendOption").options[document.getElementById("sendOption").selectedIndex].text : ''
        }
        
        var array = [letter, profiles[document.getElementById("name").value]];

        resolve(array);
    })
    
}

