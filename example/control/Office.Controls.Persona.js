var currentNode = null; 

function loadTemplate(parentElementID, templateID, isHidden)
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
                var root = document.getElementById(parentElementID);
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

    for (var i = 0; i < properties.length; i++) { 
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


function showNameOnly()
{
    loadTemplate('nameOnlyRoot', 'NameOnly', true);
}

var isCreated = true;
function showNameImage()
{
    if (isCreated)
    {
        currentNode = loadTemplate('nameOnlyRoot', 'NameImage', true);
        isCreated = false;
    }
}

function showPersonaCard()
{
    currentNode = loadTemplate('personaCardRoot', 'PersonaCard', true);
}

function showAll()
{
    showNameOnly();
    showPersonaCard();
}

function removeDetailCard()
{
   /* var root = document.getElementById("root");
    root.parentNode.removeChild(currentNode);*/
}

function createPersona()
{
    // Create NameOnly
    loadTemplate('NameOnly', true);
}

var oriID = "";
function setActiveStyle(selfObj, id)
{
    var changedClassName = " is-active";
    var parentClassName = "ms-PersonaCard-action";
    var childClassName = "ms-PersonaCard-actionDetails";

    var parentElement = document.getElementsByClassName(parentClassName);
    for (var i = 0; i < parentElement.length; i++) {
        parentElement[i].className = parentElement[i].className.replace(changedClassName, "");
    }
    var childElement = document.getElementsByClassName(childClassName);
    for (var i = 0; i < childElement.length; i++) {
        childElement[i].className = childElement[i].className.replace(changedClassName, "");
    }

    if ((selfObj.className.indexOf(changedClassName) == -1) && (oriID != id)) {
        selfObj.className += changedClassName;
        var srcElement = document.getElementById(id);
        srcElement.className += changedClassName;  
        oriID = id;
    } else {
        oriID = "";
    }
}