var tempPath = "control/templates/template.htm";
var dataProvider = sampleJsonBetter();
var personaType;
var nameImage = null;
var isShow = true;
var ips, ipc;


function showPersonaCard () {	
	var pcRoot = document.getElementById('personaCardRoot');
	// personaType = Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard;

	// Method 1:
	// var personaCard = new Office.Controls.Persona(pcRoot, dataProvider, personaType, true);
	// personaCard.loadTemplateAsync(tempPath, function (rootNode, error) {

	// });

	// Method 2:
	// Office.Controls.Persona.PersonaHelper.createPersona(pcRoot, dataProvider, personaType, callbackForPersonaCard);
	// function callbackForPersonaCard(rootNode, error) {

	// }

	// Method 3:
	ipc = Office.Controls.Persona.PersonaHelper.createPersonaCard(pcRoot, dataProvider, callbackForPersonaCard);
	function callbackForPersonaCard(rootNode, error) {

	} 
}

var interval;
function changesInLiving() {
	var pcRoot = document.getElementById('personaCardRoot');
	var keywords = ["zhongzhong li", "jonathan tang", "wenbo shi", "***REMOVED***", "jichen", "jiayuan", "abe ge"]; 

	interval = setInterval(function () {
		getAadDataDataForPersona(keywords[getRandomInt(0, 7)], pcRoot, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard, function(rootNode, error){
			});
	}, 4000);
}

function StopLiving()
{
	var pcRoot = document.getElementById('personaCardRoot');
	clearInterval(interval);
	while (pcRoot.firstChild) {
		pcRoot.removeChild(pcRoot.firstChild);
	};
	Office.Controls.Persona.PersonaHelper.createPersonaCard(pcRoot, dataProvider, function(rootNode, error) {});
}

/**
 * Returns a random integer between min (inclusive) and max (inclusive)
 * Using Math.round() will give you a non-uniform distribution!
 */
function getRandomInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

function showInlinePersona () {
	var root = document.getElementById('nameOnlyRoot');
	// personaType = Office.Controls.Persona.PersonaType.TypeEnum.NameImage;
	
	// Method 1:
	// var nameOnly = new Office.Controls.Persona(root, Office.Controls.Persona.PersonaType.TypeEnum.NameOnly, dataProvider, true);
	// nameOnly.loadTemplateAsync(tempPath, function (rootNode, error) {
 //        if (rootNode !== null) {
 //            Office.Controls.Utils.addEventListener(rootNode, 'click', function (e) {
 //            	if (nameImage == null) {
 //            		nameImage = new Office.Controls.Persona(root, personaType, dataProvider, true);
	// 	        	nameImage.loadTemplateAsync(tempPath, function (rootNode, error) {
	// 	        		isShow = false;
	// 	        	});	
 //            	} else {
 //            		nameImage.showNode(nameImage.get_rootNode(), isShow);
 //            		isShow = isShow ? false : true;
 //            	}
	// 	    });
 //        } else {
 //            // error handling
 //        }
 //    });
	
	// Method 2:
 //    Office.Controls.Persona.PersonaHelper.createPersona(root, dataProvider, personaType, callbackForNameOnly);
	// function callbackForNameOnly(rootNode, error) {
	// 	if (rootNode !== null) {
 //            Office.Controls.Utils.addEventListener(rootNode, 'click', function (e) {
 //            	if (nameImage == null) {
 //            		nameImage = new Office.Controls.Persona(root, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard, dataProvider, true);
	// 	        	nameImage.loadTemplateAsync(tempPath, function (rootNode, error) {
	// 	        		isShow = false;
	// 	        	});	
 //            	} else {
 //            		nameImage.showNode(nameImage.get_rootNode(), isShow);
 //            		isShow = isShow ? false : true;
 //            	}
	// 	    });
 //        } else {
 //            // error handling
 //        }
	// }

	// Method 3:
	ips = Office.Controls.Persona.PersonaHelper.createInlinePersona(root, dataProvider);
	event.target.disabled = true;
}

var isClickAdded = false;
function addClickEventForInlinePersona() 
{
   var showNodeQueue = [];
   var eventSpan = document.getElementById('eventDescription');
   if (!isClickAdded) {
   		isClickAdded = true; 
   		// event.target.value = "Add Click Event";
   		eventSpan.innerText = "Click Event has been added.";
   		Office.Controls.Utils.addEventListener(ips.get_rootNode(), 'click', function (e) {
			if (nameImage == null) {
				nameImage = new Office.Controls.Persona(ips.root, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard, dataProvider, true);
		    	nameImage.loadTemplateAsync(tempPath, function (rootNode, error) {
		    		showNodeQueue.push(nameImage);
		    	});	
			} else {
				if (showNodeQueue.length !== 0) {
					nameImage.showNode(nameImage.get_rootNode(), false);
					showNodeQueue.pop(nameImage);
				} else 
				{
					nameImage.showNode(nameImage.get_rootNode(), true);
					showNodeQueue.push(nameImage);
				}
			}
		});
		document.addEventListener('click', function () {
			if (event.target.tagName.toLowerCase() === "html") 
			{
				if (showNodeQueue.length !== 0) {
					nameImage.showNode(nameImage.get_rootNode(), false);
					showNodeQueue.pop(nameImage);
				}
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
var serverHost = '***REMOVED***';
function init() {
    var pageUri = window.location.href;
    pageUri = pageUri.split("?")[0];

    document.getElementById('login_user').href = 'http://' + serverHost + '/authcode?redirect_uri=' + pageUri;

    var userId = getQueryString()["userId"];
    if (typeof (userId) !== 'undefined') {
        document.getElementById('logged_user').innerText = "logged as " + userId;
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

    getAadDataDataForPersona(inputValue, root, Office.Controls.Persona.PersonaType.TypeEnum.NameImage);
}

function getAadDataDataForPersona(keyword, rootNode, personaType, callback) {
    // AAD data
    var aadDataProvider = new Office.Controls.PeopleAadDataProvider(null);
    aadDataProvider.serverHost = serverHost;
	aadDataProvider.getPrincipals(keyword, function (error, addUsers) {
	    if (addUsers !== null) {
	    	loadingImg.style.display = "none";
	        var personaObjs = Office.Controls.Persona.PersonaHelper.convertAadUsersToPersonaObjects(addUsers);
	        if (personaObjs !== null) {
	        	personaObjs.forEach(function (personaObj) {
	        		aadDataProvider.getImageAsync(personaObj.Id, function (error, imgSrc) {
                        if (imgSrc != null) {
                            personaObj.ImageUrl = imgSrc; // Get user imamge
                        }
                        // Create persona of nameimage
                        Office.Controls.Persona.PersonaHelper.createPersona(rootNode, tempPath, personaType, personaObj, true, callback);
                    });
		        });
	        }
	    } else {

	    }
	});
}

function sampleJsonBetter() {
	var persona = {
		"Id": "f567d710-09d8-458d-902f-d786234ed0d6",
		"ImageUrl": "control/images/icon.png",
		"PrimaryText": 'Cat Miao',
	    "SecondaryText": 'Software Engineer 2, DepartmentA China', // JobTitle, Department
	    "SecondaryTextShort": 'Software Engineer 2, DepartmentA China', // JobTitle, Department
	    "TertiaryText": 'BEIJING-Building1-1/12345', // Office

	    "Actions":
			{
				"Email": "catmiao@companya.com",
			    "WorkPhone": "+86(10) 12345678", 
			    "Mobile" : "+86 1861-0000-000",
			    "Skype" : "catmiao@companya.com",
			},
	    
		"Strings":
			{
				"Label":{
							"Email": "Work: ",
							"WorkPhone": "Work: ",
							"Mobile": "Mobile: ",
							"Skype": "Skype: "
						},

				"Protocol": {
								"Email": "mailto:",
								"Phone": "tel:",
								"Skype": "sip:",
							}
			}
	};
	return persona;
}