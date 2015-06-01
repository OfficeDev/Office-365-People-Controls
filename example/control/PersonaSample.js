function showAll()
{
	var dataProvider = samplePersonInfo();
	var root = document.getElementById('nameOnlyRoot');
	var nameOnly = new Office.Controls.Persona(root, 'NameOnly', dataProvider, true);
	Office.Controls.Utils.addEventListener(nameOnly.root, 'click', function (e) {
        return new Office.Controls.Persona(root, 'NameImage', dataProvider, true);
    });
	
	var pcRoot = document.getElementById('personaCardRoot');
	var personaCard = new Office.Controls.Persona(pcRoot, 'PersonaCard', dataProvider, true);
}


function samplePersonInfo()
{
    // User Profile
    var personaInfo = new Array(); 
    personaInfo["DisplayName"] = "***REMOVED*** Chen"; 
    personaInfo["JobTitle"] = "Software Engineer 2"; 
    personaInfo["ImageUrl"] = "control/images/icon.png"; 
    personaInfo["Department"] = "ASG EA China";
    personaInfo["Office"] = "BEIJING-BJW-1/12329"; 
    personaInfo["Email"] = "***REMOVED***@microsoft.com"; 
    personaInfo["SipAddress"] = "***REMOVED***@microsoft.com"; 
    personaInfo["MobilePhone"] = "+86 1861-2947-014"; 
    personaInfo["WorkPhone"] = "+86(10) 59173216"; 
    return personaInfo;
}