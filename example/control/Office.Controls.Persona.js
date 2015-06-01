(function () {
    "use strict";

    if (window.Type && window.Type.registerNamespace) {
        Type.registerNamespace('Office.Controls');
    } else {
        if (window.Office === undefined) {
            window.Office = {}; window.Office.namespace = true;
        }
        if (window.Office.Controls === undefined) {
            window.Office.Controls = {}; window.Office.Controls.namespace = true;
        }
    }

    Office.Controls.Persona = function (root, templateID, dataProvider, isHidden) {
        try {
            if (typeof root !== 'object' || typeof dataProvider !== 'object' || (Office.Controls.Utils.isNullOrUndefined(templateID))) {
                Office.Controls.Utils.errorConsole('Invalid parameters type');
                return;
            }
            
            this.root = root;
            this.templateID = templateID;
            this.dataProvider = dataProvider;
            this.isHidden = isHidden;
            this.currentNode = null;

            this.loadTemplate();
        } catch (ex) {
            throw ex;
        }
    };

    Office.Controls.Persona.prototype = {
        delaySearchInterval: 300,
        onError: null,
        dataProvider: null,
        showErrors: true,        
        errorMessageElement: null,
        alertDiv: null,
        currentToken: null,
        widthSet: false,
        hasErrors: false,
        errorDisplayed: null,
        isCreated: true,
        oriID: "",

        getErrorDisplayed: function () {
            return this.errorDisplayed;
        },

        loadTemplate: function ()
        {
            var self = this;
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
                        
                        var templateNode = xmlDoc.getElementById(self.templateID);
                        var templateElement = document.createElement("div");
                        self.hiddenNode(templateElement, self.isHidden);

                        // Do data binding
                        templateElement.innerHTML = self.bindData(templateNode.innerHTML);

                        self.root.appendChild(templateElement);
                        self.currentNode = templateElement;
                    }
                }
                else
                {
                    // error message
                }
                
            };
            xmlhttp.send();
        },

        bindData: function (htmlStr)
        {
            var regExp = /\$\{([^\}\{]+)\}/g;
            var resultStr = htmlStr;
            // Get the data
            var personaInfo = this.dataProvider;

            // Get the property names
            var properties = resultStr.match(regExp);

            for (var i = 0; i < properties.length; i++) { 
                var propertyValue = personaInfo[properties[i].substring(2, properties[i].length - 1)]
                resultStr = resultStr.replace(properties[i], propertyValue);
            }

            return resultStr;
        },

        hiddenNode: function (node, isHidden)
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
        },

        removeDetailCard: function ()
        {
           /* var root = document.getElementById("root");
            root.parentNode.removeChild(currentNode);*/
        },

        setActiveStyle: function (event, childID)
        {
            // Get the element triggers the event
            var e = event || window.event;
            var targetElement = e.target || e.srcElement;

            var changedClassName = " is-active";
            var parentClassName = Office.Controls.PersonaConstants.ClassName_PersonaCardAction;
            var childClassName = PersonaConstants.ClassName_PersonaCardActionDetails;

            var parentElements = document.getElementsByClassName(parentClassName);
            for (var i = 0; i < parentElements.length; i++) {
                parentElements[i].className = parentElements[i].className.replace(changedClassName, "");
            }
            var childElements = document.getElementsByClassName(childClassName);
            for (var i = 0; i < childElements.length; i++) {
                childElements[i].className = childElements[i].className.replace(changedClassName, "");
            }

            if ((targetElement.className.indexOf(changedClassName) == -1) || (oriID != childID)) {
                targetElement.className += changedClassName;
                var srcElement = document.getElementById(childID);
                srcElement.className += changedClassName;  
                this.oriID = childID;
            } else {
                this.oriID = "";
            }
        },
    };

    Office.Controls.PersonaResources = function () {
    };

    Office.Controls.PersonaConstants = function () {
    };

    if (Office.Controls.Persona.registerClass) { Office.Controls.Persona.registerClass('Office.Controls.Persona'); }
    if (Office.Controls.PersonaConstants.registerClass) { Office.Controls.PersonaConstants.registerClass('Office.Controls.PersonaConstants'); }
    if (Office.Controls.PersonaResources.registerClass) { Office.Controls.PersonaResources.registerClass('Office.Controls.PersonaResources'); }
    Office.Controls.Persona.res = {};
    Office.Controls.PersonaConstants.ClassName_PersonaCardAction = "ms-PersonaCard-action";
    Office.Controls.PersonaConstants.ClassName_PersonaCardActionDetails = "ms-PersonaCard-actionDetails";
    Office.Controls.PersonaResources.PersonaName = 'Persona';
})();