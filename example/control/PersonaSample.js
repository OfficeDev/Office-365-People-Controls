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


function samplePersonInfo() {
    // User Profile
    return displayInfo = {
            ImageUrl: 'control/images/icon.png',
            PrimaryText: '***REMOVED*** Chen',
            SecondaryText: 'Software Engineer 2, ASG EA China', // JobTitle, Department
            TertiaryText: 'BEIJING-BJW-1/12329', // Office
            Email: '***REMOVED***@microsoft.com',
            SipAddress: '***REMOVED***@microsoft.com',
            MobilePhone: '+86 1861-2947-014',
            WorkPhone: '+86(10) 59173216',
            Id: '***REMOVED***',
        };
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
							"Work": { "Label": "Work: ", "Protocol": "mailto:", "Value": "***REMOVED***@microsoft.com" }
						},

			    "Phone": 
			    		{
							"Work": { "Label": "Work: ", "Protocol": "mailto:", "Value": "+86(10) 59173216" },
							"Mobile": { "Label": "Mobile: ", "Protocol": "mailto:", "Value": "+86 1861-2947-014" }
						},

	    		"Chat": 
			    		{
							"Lync": { "Label": "Lync: ", "Protocol": "sip:", "Value": "***REMOVED***@microsoft.com" },
						}
			}
	};
	return persona;
}