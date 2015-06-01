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

    Office.Controls.Persona = function (parentElementID, templateID, isHidden) {
        try {
            if ((Office.Controls.Utils.isNullOrUndefined(parentElementID)) || (Office.Controls.Utils.isNullOrUndefined(templateID))) {
                Office.Controls.Utils.errorConsole('Invalid parameters type');
                return;
            }
            this.parentElementID = parentElementID;
            this.templateID = templateID;
            this.isHidden = isHidden;
            this.currentNode = null;
            this.root = document.getElementById(parentElementID);            

            this.loadTemplate(this.parentElementID, this.templateID, this.isHidden);
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

        getDisplayInfoAsync: function (userInfoHandler, userEmail) {
        },

        renderControl: function (inputName) {
            // this.root.innerHTML = Office.Controls.PersonaTemplates.generateControlTemplate(inputName, this.allowMultiple, this.inputHint);
            // if (this.root.className.length > 0) {
            //     this.root.className += ' ';
            // }
            // this.root.className += 'office office-peoplepicker';
            // this.actualRoot = this.root.querySelector('div.ms-PeoplePicker');
            // var self = this;
            // Office.Controls.Utils.addEventListener(this.actualRoot, 'click', function (e) {
            //     return self.onPickerClick(e);
            // });
            // this.inputData = this.actualRoot.querySelector('input[type=\"hidden\"]');
            // this.textInput = this.actualRoot.querySelector('input[type=\"text\"]');
            // this.defaultText = this.actualRoot.querySelector('span.office-peoplepicker-default');
            // this.resolvedListRoot = this.actualRoot.querySelector('div.office-peoplepicker-recordList');
            // this.autofillElement = this.actualRoot.querySelector('.ms-PeoplePicker-results');
            // this.alertDiv = this.actualRoot.querySelector('.office-peoplepicker-alert');
            // Office.Controls.Utils.addEventListener(this.textInput, 'focus', function (e) {
            //     return self.onInputFocus(e);
            // });
            // Office.Controls.Utils.addEventListener(this.textInput, 'blur', function (e) {
            //     return self.onInputBlur(e);
            // });
            // Office.Controls.Utils.addEventListener(this.textInput, 'keydown', function (e) {
            //     return self.onInputKeyDown(e);
            // });
            // Office.Controls.Utils.addEventListener(this.textInput, 'input', function (e) {
            //     return self.onInput(e);
            // });
            // Office.Controls.Utils.addEventListener(window.self, 'resize', function (e) {
            //     return self.onResize(e);
            // });
            // this.toggleDefaultText();
            // if (!Office.Controls.Utils.isNullOrUndefined(this.inputTabindex)) {
            //     this.textInput.setAttribute('tabindex', this.inputTabindex);
            // }
        },

        loadTemplate: function (parentElementID, templateID, isHidden)
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
                        
                        var templateNode = xmlDoc.getElementById(templateID);
                        var templateElement = document.createElement("div");
                        self.hiddenNode(templateElement, isHidden);

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

        samplePersonInfo: function ()
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
        },

        bindData: function (htmlStr)
        {
            var regExp = /\$\{([^\}\{]+)\}/g;
            var resultStr = htmlStr;
            // Get the data
            var personaInfo = this.samplePersonInfo();

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

        showNameOnly: function ()
        {
            loadTemplate('nameOnlyRoot', 'NameOnly', true);
            var self = this;
            Office.Controls.Utils.addEventListener(this.root, 'click', function (e) {
                return self.showNameImage();
            });
        },

        showPersonaCard: function ()
        {
            currentNode = loadTemplate('personaCardRoot', 'PersonaCard', true);
        },

        removeDetailCard: function ()
        {
           /* var root = document.getElementById("root");
            root.parentNode.removeChild(currentNode);*/
        },

        createPersona: function ()
        {
            // Create NameOnly
            loadTemplate('NameOnly', true);
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
    // if (Office.Controls.PersonaTemplates.registerClass) { Office.Controls.PersonaTemplates.registerClass('Office.Controls.PersonaTemplates'); }
    if (Office.Controls.PersonaConstants.registerClass) { Office.Controls.PersonaConstants.registerClass('Office.Controls.PersonaConstants'); }
    if (Office.Controls.PersonaResources.registerClass) { Office.Controls.PersonaResources.registerClass('Office.Controls.PersonaResources'); }
    Office.Controls.Persona.res = {};
    Office.Controls.PersonaConstants.ClassName_PersonaCardAction = "ms-PersonaCard-action";
    Office.Controls.PersonaConstants.ClassName_PersonaCardActionDetails = "ms-PersonaCard-actionDetails";
    Office.Controls.PersonaResources.PersonaName = 'Persona';
})();