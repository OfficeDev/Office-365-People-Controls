var currentNode = null; 

function loadTemplate(templateID, isHidden)
{
    var xmlhttp = new XMLHttpRequest();
    xmlhttp.open("GET", "control/templates/template.htm", true);
    xmlhttp.onreadystatechange = function() {
        // if (this.readyState !== 4) return;
        // if (this.status !== 200) return; // or whatever error handling you want
        if (this.readyState === 4)
        {
            if (this.status === 200)
            {
                var parser, xmlDoc
                if (window.DOMParser)
                {
                   parser = new DOMParser();
                   xmlDoc = parser.parseFromString(this.responseText,"text/xml");
                }
                else // code for IE
                {
                   xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
                   xmlDoc.async = false;
                   xmlDoc.loadXML(this.responseText); 
                }  
                
                var templateNode = xmlDoc.getElementById(templateID);
                var root = document.getElementById("root");
                var templateElement = document.createElement("div");
                hiddenNode(templateElement, isHidden);

                // Do data binding
                templateElement.innerHTML = bindData(templateNode.innerHTML);

                root.appendChild(templateElement);
                currentNode = templateElement;
            }
        }
        else
        {
            // error message
        }
        
    };
    xmlhttp.send();
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

function bindData(htmlStr)
{
    var regExp = /\$\{([^\}\{]+)\}/g;
    var resultStr = htmlStr;
    // Get the data
    var personaInfo = samplePersonInfo();

    // Get the property names
    var properties = resultStr.match(regExp);

    for (i = 0; i < properties.length; i++) { 
        var propertyValue = personaInfo[properties[i].substring(2, properties[i].length - 1)]
        resultStr = resultStr.replace(properties[i], propertyValue);
    }

    return resultStr;
}

function hiddenNode(node, isHidden)
{
    if (isHidden)
    {
        node.style.visibility = "";
        node.style.display = "";
    }
    else
    {
        node.style.visibility = "hidden";
        node.style.display = "none";
    }
}

var isCreated = true;
function showDetailCard()
{
    if (isCreated)
    {
        currentNode = loadTemplate('DetailCard', true);
        isCreated = false;
    }
}

function removeDetailCard()
{
    var root = document.getElementById("root");
    root.parentNode.removeChild(currentNode);
}

function createPersona()
{
    // Create NameOnly
    loadTemplate('NameOnly', true);
}