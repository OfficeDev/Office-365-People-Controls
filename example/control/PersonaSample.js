function showAll()
{	
	var tempPath = "control/templates/template.htm";
	var dataProvider = samplePersonInfo();
	var root = document.getElementById('nameOnlyRoot');
	var nameOnly = new Office.Controls.Persona(root, 'NameOnly', dataProvider, true);
	var nameImage = null;
	var isShow = true;

	nameOnly.loadTemplateAsync(tempPath, function (rootNode, error) {
        if (rootNode !== null) {
            Office.Controls.Utils.addEventListener(rootNode, 'click', function (e) {
            	if (nameImage == null) {
            		nameImage = new Office.Controls.Persona(root, 'NameImage', dataProvider, true);
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
    });
	
	var pcRoot = document.getElementById('personaCardRoot');
	var personaCard = new Office.Controls.Persona(pcRoot, 'PersonaCard', dataProvider, true);
	personaCard.loadTemplateAsync(tempPath, function (rootNode, error) {

	});
}


function samplePersonInfo()
{
    // User Profile
    // return personaInfo = {
    //     DisplayName: '***REMOVED*** Chen',
    //     JobTitle: 'Software Engineer 2',
    //     ImageUrl: 'control/images/icon.png',
    //     Department: 'ASG EA China',
    //     Office: 'BEIJING-BJW-1/12329',
    //     Email: '***REMOVED***@microsoft.com',
    //     SipAddress: '***REMOVED***@microsoft.com',
    //     MobilePhone: '+86 1861-2947-014',
    //     WorkPhone: '+86(10) 59173216',
    //     PersonId: '***REMOVED***',
    // };

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