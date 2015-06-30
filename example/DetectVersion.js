var loadjscssfile = function (filename, filetype) {
    if (filetype === "js") {
        var fileref = document.createElement('script');
        fileref.setAttribute("src", filename);
    } else if (filetype === "css") {
        var fileref = document.createElement('link');
        fileref.setAttribute("rel", "stylesheet");
        fileref.setAttribute("type", "text/css");
        fileref.setAttribute("href", filename);
        fileref.setAttribute("media", "all");
    }
    if (typeof fileref !== "undefined") {
        document.getElementsByTagName("head")[0].appendChild(fileref);
    }
}

var getArgs = function () {
    var args = new Object();
    var query = window.location.search.substring(1);
    var pairs = query.split("&");
    for (var i = 0; i < pairs.length; i++) {
        var pos = pairs[i].indexOf('=');
        if (pos == -1) continue;
        var argname = pairs[i].substring(0, pos);
        var value = pairs[i].substring(pos + 1);
        value = decodeURIComponent(value);
        args[argname] = value;
    }
    return args;
}

var peoplePickerCSS = "control/Office.Controls.PeoplePicker.min.css";
var peoplePicerBaseJS = "control/Office.Controls.Base.min.js";
var peoplePicerJS = "control/Office.Controls.PeoplePicker.min.js";
var peoplePicerProviderJS = "control/Office.Controls.PeopleAadDataProvider.min.js";

// Get URL Parameter and load debug/minify version
// debug=1 -- debug version; debug=0 --minify version
var args = getArgs();
var isdebug = args.debug || "0";

if (isdebug === "1") {
    peoplePickerCSS = "control/Office.Controls.PeoplePicker.css";
    peoplePicerBaseJS = "control/Office.Controls.Base.js";
    peoplePicerJS = "control/Office.Controls.PeoplePicker.js";
    peoplePicerProviderJS = "control/Office.Controls.PeopleAadDataProvider.js";
}

loadjscssfile(peoplePickerCSS, "css");
loadjscssfile(peoplePicerBaseJS, "js");
loadjscssfile(peoplePicerJS, "js");
loadjscssfile(peoplePicerProviderJS, "js");

