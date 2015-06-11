var tempPath = "control/templates/template.htm";
var dataProvider = sampleJsonBetter();
var personaType;
var nameImage = null;
var isShow = true;
var ips, ipc;

function showPersonaCard () {	
	var pcRoot = document.getElementById('personaCardRoot');
	personaType = Office.Controls.Persona.PersonaHelper.getPersonaType().PersonaCard;

	// Method 1:
	// var personaCard = new Office.Controls.Persona(pcRoot, 'personacard', dataProvider, true);
	// personaCard.loadTemplateAsync(tempPath, function (rootNode, error) {

	// });

	// Method 2:
	// Office.Controls.Persona.PersonaHelper.createPersona(pcRoot, tempPath, personaType, dataProvider, true, callbackForPersonaCard);
	// function callbackForPersonaCard(rootNode, error) {

	// }

	// Method 3:
	ipc = Office.Controls.Persona.PersonaHelper.createPersonaCard(pcRoot, tempPath, dataProvider, callbackForPersonaCard);
	function callbackForPersonaCard(rootNode, error) {

	}
}

function showInlinePersona () {
	var root = document.getElementById('nameOnlyRoot');
	personaType = Office.Controls.Persona.PersonaHelper.getPersonaType().NameImage;
	
	// Method 1:
	// var nameOnly = new Office.Controls.Persona(root, 'nameonly', dataProvider, true);
	// nameOnly.loadTemplateAsync(tempPath, function (rootNode, error) {
 //        if (rootNode !== null) {
 //            Office.Controls.Utils.addEventListener(rootNode, 'click', function (e) {
 //            	if (nameImage == null) {
 //            		nameImage = new Office.Controls.Persona(root, 'nameimage', dataProvider, true);
	// 	        	nameImage.loadTemplateAsync(tempPath, function (rootNode, error) {
	// 	        		isShow = false;
	// 	        	});	
 //            	} else {
 //            		nameImage.hiddenNode(nameImage.get_rootNode(), isShow);
 //            		isShow = isShow ? false : true;
 //            	}
	// 	    });
 //        } else {
 //            // error handling
 //        }
 //    });
	
	// Method 2:
 //    Office.Controls.Persona.PersonaHelper.createPersona(root, tempPath, personaType, dataProvider, true, callbackForNameOnly);
	// function callbackForNameOnly(rootNode, error) {
	// 	if (rootNode !== null) {
 //            Office.Controls.Utils.addEventListener(rootNode, 'click', function (e) {
 //            	if (nameImage == null) {
 //            		nameImage = new Office.Controls.Persona(root, Office.Controls.Persona.PersonaHelper.getPersonaType().PersonaCard, dataProvider, true);
	// 	        	nameImage.loadTemplateAsync(tempPath, function (rootNode, error) {
	// 	        		isShow = false;
	// 	        	});	
 //            	} else {
 //            		nameImage.hiddenNode(nameImage.get_rootNode(), isShow);
 //            		isShow = isShow ? false : true;
 //            	}
	// 	    });
 //        } else {
 //            // error handling
 //        }
	// }

	// Method 3:
	ips = Office.Controls.Persona.PersonaHelper.createInlinePersona(root, tempPath, dataProvider);
	event.target.disabled = true;
}

var isClickAdded = false;
function addClickEventForInlinePersona() 
{
   var eventSpan = document.getElementById('eventDescription');
   if (!isClickAdded) {
   		isClickAdded = true; 
   		// event.target.value = "Add Click Event";
   		eventSpan.innerText = "Click Event has been added.";
   		Office.Controls.Utils.addEventListener(ips.get_rootNode(), 'click', function (e) {
			if (nameImage == null) {
				nameImage = new Office.Controls.Persona(ips.root, Office.Controls.Persona.PersonaHelper.getPersonaType().PersonaCard, dataProvider, true);
		    	nameImage.loadTemplateAsync(tempPath, function (rootNode, error) {
		    		isShow = false;
		    	});	
			} else {
				nameImage.hiddenNode(nameImage.get_rootNode(), isShow);
				isShow = isShow ? false : true;
			}
		});
   } else {
   		isClickAdded = false;
   		eventSpan.innerText = "Click Event has been removed.";
   		// event.target.value = "Remove Click Event";
   		Office.Controls.Utils.removeEventListener(ips.get_rootNode(), 'click', function (e) {});
   }
}

/**
 * Integrate with AAD Data
 */
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

function isCreatePersona() {
	if(event.keyCode === 13){
       createPersonaWithAadData();
    }
}

function createPersonaWithAadData() {
	var input = document.getElementById('keyword');
	var inputValue = input.value.trim();
    if (inputValue === "")
    {
    	alert('please input the keyword!');
    	return;
    }

	var root = document.getElementById('aadUserRoot');
	while (root.firstChild) {
		root.removeChild(root.firstChild);
	};

    var loadingImg = document.getElementById('loadingImg');
    loadingImg.style.display = "";
    

    serverHost = '***REMOVED***';

    var pageUri = window.location.href;
    pageUri = pageUri.split("?")[0];

    document.getElementById('login_user').href = 'http://' + serverHost + '/authcode?redirect_uri=' + pageUri;

    var userId = getQueryString()["userId"];
    if (typeof (userId) !== 'undefined') {
        document.getElementById('logged_user').innerText = "logged as " + userId;
    }

    // AAD data
    var aadDataProvider = new Office.Controls.PeopleAadDataProvider();
    aadDataProvider.serverHost = serverHost;
	aadDataProvider.getPrincipals(inputValue, function (error, addUsers) {
	    if (addUsers !== null) {
	    	loadingImg.style.display = "none";
	        var personaObjs = Office.Controls.Persona.PersonaHelper.convertAadUsersToPersonaObjects(addUsers);
	        if (personaObjs !== null) {
	        	personaObjs.forEach(function (personaObj) {
	        		personaObj.Main.ImageUrl = "control/images/doughboy.png";
		            Office.Controls.Persona.PersonaHelper.createInlinePersona(root, tempPath, personaObj);
		        });
	        }

	    } else {

	    }
	});
}

function sampleJsonBetter() {
	var persona = {
		"Id": "***REMOVED***",
		"Main":
			{
				"ImageUrl": "control/images/icon.png",
				"PrimaryText": '***REMOVED*** Chen',
	            "SecondaryText": 'Software Engineer 2, ASG EA China', // JobTitle, Department
	            "TertiaryText": 'BEIJING-BJW-1/12329', // Office
			},

		"Action":
			{
				"Email":{
							"Protocol": "mailto:",
							"Work": { "Label": "Work: ", "Value": "***REMOVED***@microsoft.com" }
						},

			    "Phone": 
			    		{
							"Protocol": "tel:",
							"Work": { "Label": "Work: ", "Value": "+86(10) 59173216" },
							"Mobile": { "Label": "Mobile: ", "Value": "+86 1861-2947-014" }
						},

	    		"Chat": 
			    		{
							"Protocol": "sip:",
							"Skype": { "Label": "Skype: ", "Value": "***REMOVED***@microsoft.com" },
						}
			}
	};
	return persona;
}