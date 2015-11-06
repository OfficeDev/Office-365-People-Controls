var dataProvider = sampleAADObj();
var nameImage = null;
var ips, ipc;

function showPersonaCard () {	
	var pcRoot = document.getElementById('personaCardRoot');
	ipc = Office.Controls.Persona.PersonaHelper.createPersonaCard(pcRoot, dataProvider);
}

function showInlinePersona () {
	var root = document.getElementById('nameOnlyRoot');
	ips = Office.Controls.Persona.PersonaHelper.createInlinePersona(root, dataProvider);
}

var isClickAdded = false;
function addClickEventForInlinePersona() 
{
   var showNodeQueue = [];
   var eventSpan = document.getElementById('eventDescription');
   if (!isClickAdded) {
   		isClickAdded = true; 
   		eventSpan.innerText = "Click Event has been added.";
   		Office.Controls.Utils.addEventListener(ips.get_rootNode(), 'click', function (e) {
			if (nameImage == null) {
				nameImage = Office.Controls.Persona.PersonaHelper.createPersonaCard(ips.root, dataProvider);
		    	showNodeQueue.push(nameImage);
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

		Office.Controls.Utils.addEventListener(document, 'click', function () {
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
   		Office.Controls.Utils.removeEventListener(ips.get_rootNode(), 'click', function (e) {});
   }
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

    getAadDataDataForPersona(inputValue, root);
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

function sampleAADObj() {
    var persona = {
		"id": "f567d710-09d8-458d-902f-d786234ed0d6",
		"displayName": 'Jerry Anderson',
	    "department": 'DepartmentA China',
	    "jobTitle": 'Software Engineer',
	    "office": 'BEIJING-Building1-1/12345', // Office
		"mail": "jerryanderson@companya.com",
		"workPhone": "+86(10) 12345678", 
		"mobile" : "+86 1861-0000-000",
		"sipAddress" : "jerryanderson@companya.com"
	};
	return persona;
}