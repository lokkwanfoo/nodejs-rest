Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.

        $("#close").click(function () {
            postData();
            Office.context.ui.messageParent(JSON.stringify(letter));
            window.close();
        });

        $("#saveLetter").click(function () {
            postData();
        });

    });
}

var accessToken = localStorage.getItem("accessToken");
var letter;

function postData() {

    letter = {
        "nameaddress":  document.getElementById("nameaddress").value,
        "yourReference": document.getElementById("yourReference").value,
        "ourReference": document.getElementById("ourReference").value,
        "subject": document.getElementById("subject").value,
        "name": document.getElementById("name").value
    }
    
}

