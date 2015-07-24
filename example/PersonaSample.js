/**
 * Integrate with AAD Data
 */
var serverHost = '***REMOVED***';
function init() {
    var pageUri = window.location.href;
    pageUri = pageUri.split("?")[0];

    document.getElementById('login_user').href = 'http://' + serverHost + '/authcode?redirect_uri=' + pageUri;

    var userId = getQueryString()["userId"];
    if (typeof (userId) !== 'undefined') {
        document.getElementById('logged_user').textContent = "logged as " + userId;
    }
}

function getQueryString() {
    var result = {}, queryString = location.search.slice(1),
        re = /([^&=]+)=([^&]*)/g, m;

    while (m = re.exec(queryString)) {
        result[decodeURIComponent(m[1])] = decodeURIComponent(m[2]);
    }

    return result;
}

function getAadDataDataForPersona(keyword, rootNode) {
    // AAD data
    var aadDataProvider = new AadDataProvider(null);
    aadDataProvider.serverHost = serverHost;
	aadDataProvider.searchPeopleAsync(keyword, function (error, personObjs) {
	    if (personObjs !== null) {
	    	loadingImg.style.display = "none";
	    	personObjs.forEach(function (person) {
        		aadDataProvider.getImageAsync(person.id, function (error, imgSrc) {
                    if (imgSrc != null) {
                        person.imgSrc = imgSrc; // Get user imamge
                    }
                    // Create persona of nameimage
                    Office.Controls.Persona.PersonaHelper.createInlinePersona(rootNode, person);
                });
	        });
	    } else {
	    }
	});
}