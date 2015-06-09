var tempPath = "control/templates/template.htm";
var dataProvider = sampleJsonBetter();
var personaType = Office.Controls.Persona.PersonaHelper.getPersonaType().NameOnly;
var nameImage = null;
var isShow = true;

function showPersonaCard()
{	
	var pcRoot = document.getElementById('personaCardRoot');
	// Method 1:
	// var personaCard = new Office.Controls.Persona(pcRoot, 'personacard', dataProvider, true);
	// personaCard.loadTemplateAsync(tempPath, function (rootNode, error) {

	// });

	// Method 2:
	personaType = Office.Controls.Persona.PersonaHelper.getPersonaType().PersonaCard;
	Office.Controls.Persona.PersonaHelper.createPersona(pcRoot, tempPath, personaType, dataProvider, true, callbackForPersonaCard);
	function callbackForPersonaCard(rootNode, error) {

	}
}

function showNameOnly () {
	var root = document.getElementById('nameOnlyRoot');
	personaType = Office.Controls.Persona.PersonaHelper.getPersonaType().NameOnly;
	
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
    Office.Controls.Persona.PersonaHelper.createPersona(root, tempPath, personaType, dataProvider, true, callbackForNameOnly);
	function callbackForNameOnly(rootNode, error) {
		if (rootNode !== null) {
            Office.Controls.Utils.addEventListener(rootNode, 'click', function (e) {
            	if (nameImage == null) {
            		nameImage = new Office.Controls.Persona(root, Office.Controls.Persona.PersonaHelper.getPersonaType().NameImage, dataProvider, true);
		        	nameImage.loadTemplateAsync(tempPath, function (rootNode, error) {
		        		isShow = false;
		        	});	
            	} else {
            		nameImage.hiddenNode(nameImage.get_rootNode(), isShow);
            		isShow = isShow ? false : true;
            	}
		    });
        } else {
            // error handling
        }
	}
}


function samplePersonInfo()
{
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

function sampleJsonBetter()
{
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