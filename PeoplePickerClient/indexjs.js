
       var serverHost = 'yihcaow1001:3000';
if (window.location.href.indexOf('azurewebsites.net') != -1) {
    serverHost = '***REMOVED***';
}
var pageUri = window.location.href;
pageUri = pageUri.split("?")[0];

document.getElementById('login_user').href = 'http://' + serverHost + '/authcode?redirect_uri=' + pageUri;

var userId = getQueryString()["userId"];
if (typeof(userId) !== 'undefined') {
    document.getElementById('logged_user').innerText = "logged as " + userId;
}

// Mock data

var mockDataProvider = new MockDataProvider();
var pp = new Office.Controls.PeoplePicker(document.getElementById('ppc_mock'), mockDataProvider, {});
document.getElementById('ppc_mock').datapp = pp;

function checkbox_ignore_click(cb) {
    mockDataProvider.ignoreKeyword = cb.checked;
}

// AAD data

var aadDataProvider = new AadDataProvider();
aadDataProvider.serverHost = serverHost;

params = new Object();
params.allowMultipleSelections = false;
params.enableCache = true;
params.showValidationErrors = true;
new Office.Controls.PeoplePicker(document.getElementById('ppc_single'), aadDataProvider, params);

params.allowMultipleSelections = true;
params.startSearchCharLength = 1;
params.inputHint = "Try to select multiple records...";
params.showValidationErrors = false;
params.onError = function (control, validationError) {
    if (validationError.errorName == 'ServerProblem') {
        document.getElementById('ppc_multiple_error').innerHTML = "<pre>" + aadDataProvider.lastErrorMessage + "</pre>";
    } else {
        document.getElementById('ppc_multiple_error').innerHTML = "<pre>" + validationError.localizedErrorMessage + "</pre>";
    }
};
params.onAdded = params.onRemoved = function (control) {
    document.getElementById('ppc_multiple_error').innerHTML = "";
    var people = 'Added people: ';
    control.getAddedPeople().forEach(
        function (e) {
            people += '<p>{' + e.DisplayName + ', id=' + e.PersonId + '}</p>';
        });
    document.getElementById('ppc_multiple_people').innerHTML = "<pre>" + people + "</pre>";
}
pp = new Office.Controls.PeoplePicker(document.getElementById('ppc_multiple'), aadDataProvider, params);
document.getElementById('ppc_multiple').datapp = pp;

document.getElementById('login_instructions').style.display = 'none';
        
function showLoginInstructions() {
    var display = document.getElementById('login_instructions').style.display;
    document.getElementById('login_instructions').style.display = (display === 'none') ? 'inherit' : 'none';
}

function getQueryString() {
    var result = {}, queryString = location.search.slice(1),
        re = /([^&=]+)=([^&]*)/g, m;

    while (m = re.exec(queryString)) {
        result[decodeURIComponent(m[1])] = decodeURIComponent(m[2]);
    }

    return result;
}
