/**
 * Integrate with AAD Data
 */
var authContext = null;
function init() {

    // AAD authentication

    window.config = {
        instance: 'https://login.microsoftonline.com/',
        clientId: '<ClientId>', // Get this from Azure app you created
        postLogoutRedirectUri: window.location,
        cacheLocation: 'localStorage', // enable this for IE, as sessionStorage does not work for localhost.
    };
    authContext = new AuthenticationContext(config);

    // Check For & Handle Redirect From AAD After Login
    var isCallback = authContext.isCallback(window.location.hash);
    authContext.handleWindowCallback();

    var user = authContext.getCachedUser();
    if (user) {
        document.getElementById('logged_user').textContent = "logged as " + user.userName;
        document.getElementById('login_user').textContent = 'Logout';
    }

    document.getElementById('login_user').addEventListener("click", function () {
        if (user) {
            authContext.logOut();
        } else {
            authContext.login();
        }
    });
}

function getAadDataDataForPersona(keyword, rootNode) {
    // AAD data
    var aadDataProvider = new Office.Controls.PeopleAadDataProvider(authContext);
    aadDataProvider.getPrincipals(keyword, function (error, personObjs) {
        if (personObjs !== null) {
            loadingStr.style.display = "none";
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