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

    /*
    *  The format of DataProvider:
    *  {
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
                                "Work": { "Label": "Work: ", "Protocol": "tel:", "Value": "+86(10) 59173216" },
                                "Mobile": { "Label": "Mobile: ", "Protocol": "tel:", "Value": "+86 1861-2947-014" }
                            },
                    "Chat": 
                            {
                                "Lync": { "Label": "Lync: ", "Protocol": "sip:", "Value": "***REMOVED***@microsoft.com" },
                            }
                }
        };
    */
    Office.Controls.Persona = function (root, templateID, dataProvider, isHidden) {
        if (typeof root !== 'object' || typeof dataProvider !== 'object' || (Office.Controls.Utils.isNullOrUndefined(templateID))) {
                Office.Controls.Utils.errorConsole('Invalid parameters type');
                return;
        }
            
        this.root = root;
        this.templateID = templateID;
        this.dataProvider = dataProvider;
        this.isHidden = isHidden;
    };

    Office.Controls.Persona.prototype = {
        onError: null,
        rootNode: null,
        mainNode: null,
        actionNodes: null,
        actionDetailNodes: null,

        oriID: "",

        get_rootNode: function() {
            return this.rootNode;
        },

        get_mainNode: function() {
            return this.mainNode;
        },

        get_actionNodes: function() {
            return this.actionNodes;
        },

        get_actionDetailNodes: function() {
            return this.actionDetailNodes;
        },

        /**
         * Load template file from the give path
         * @templatePath  {string}
         * @callback  {Function}
         * @return {null}
         */
        loadTemplateAsync: function (templatePath, callback) {
            var self = this;
            var cachedTemplate = Office.Controls.Persona.PersonaHelper.getCachedTemplate(templatePath);

            if (cachedTemplate !== null)
            {
               self.parseTemplate(cachedTemplate);
               callback(self.rootNode, null);
               return;
            } 

            var xmlhttp = new XMLHttpRequest();
            xmlhttp.open("GET", templatePath, true);
            xmlhttp.onreadystatechange = function() {
                // if (this.readyState !== 4) return;
                // if (this.status !== 200) return; // or whatever error handling you want
                if (this.readyState === 4) {
                    if (this.status === 200) {
                        var parser, xmlDoc;
                        if (window.DOMParser) {
                           parser = new DOMParser();
                           xmlDoc = parser.parseFromString(this.responseText,"text/xml");
                        } else { // code for < IE9
                           xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
                           xmlDoc.async = false;
                           xmlDoc.loadXML(this.responseText); 
                        }  

                        // Save xmlDoc to cache
                        Office.Controls.Persona.PersonaHelper.setCachedTemplate(templatePath, xmlDoc);
                        self.parseTemplate(xmlDoc);
                        callback(self.rootNode, null); 
                    }
                }
            };
            xmlhttp.send();
        },

        /**
         * Parse the persona content loading from template that includes 3 parts:
         *     1. Main: It's a detail card
         *     2. Action bar: It includes the action icons and the click event listener is also attached to each icon.
         *     3. The detail content of each Action icon: When click the icon, the detail shows up.
         * @param  {[DocumentElment} xmlDoc The document loading from template
         * @return {null}
         */
        parseTemplate: function (xmlDoc) {
            try {
                var templateNode = xmlDoc.getElementById(this.templateID);
                var templateElement = document.createElement("div");
                this.hiddenNode(templateElement, this.isHidden);

                // Do data binding
                templateElement.innerHTML = this.bindData(templateNode.innerHTML);
                if ((Office.Controls.Utils.isNullOrUndefined(templateElement))) {
                    Office.Controls.Utils.errorConsole('Fail to get persona document');
                    return;
                }
                this.root.appendChild(templateElement);
                this.rootNode = templateElement;

                // Get main node
                this.mainNode = templateElement.getElementsByClassName(Office.Controls.PersonaConstants.SectionTag_Main);
                if (this.mainNode === null) {
                    this.mainNode = this.rootNode;
                } else {
                    // Get action nodes
                    this.actionNodes = templateElement.getElementsByClassName(Office.Controls.PersonaConstants.SectionTag_Action);
                    if (this.actionNodes !== null) {
                        // Get actiondetail nodes
                        this.actionDetailNodes = templateElement.getElementsByClassName(Office.Controls.PersonaConstants.SectionTag_ActionDetail);
                        // Add click event to show the action detail node
                        var node = null;
                        var self = this;
                        for (var i = 0; i < self.actionNodes.length; i++) {
                            if (self.actionNodes[i] !== null) {
                                node = self.actionNodes[i];
                                Office.Controls.Utils.addEventListener(node, 'click', function (e) {
                                    self.setActiveStyle(event);
                                });
                            }
                        }
                    }
                }
            } catch (ex) {
                throw ex;
            }
        },

        /**
         * Bind data to template
         * @htmlStr  {string}
         * @return {string}
         */
        bindData: function (htmlStr) {
            var regExp = /\$\{([^\}\{]+)\}/g;
            var resultStr = htmlStr;

            // Get the property names
            var properties = resultStr.match(regExp);
            var propertyName, propertyValue;
            var self = this;
            for (var i = 0; i < properties.length; i++) { 
                propertyName = properties[i].substring(2, properties[i].length - 1);
                propertyValue = self.getValueFromJSONObject(propertyName);
                resultStr = resultStr.replace(properties[i], propertyValue);
            }

            return resultStr;
        },

        /**
         * Parse the json object to get the corresponding value
         * @objectName  {string}
         * @return {object}
         */
        getValueFromJSONObject: function (objectName) {
            return Office.Controls.Utils.getObjectFromJSONObjectName(this.dataProvider, objectName);
        },

        hiddenNode: function (node, isHidden) {
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

        removeDetailCard: function () {
        },

        /**
         * [setActiveStyle description]
         * @param {[type]}
         */
        setActiveStyle: function (event) {
            // Get the element triggers the event
            var e = event || window.event;
            // var currentNode = e.target || e.srcElement;
            var currentNode = e.currentTarget;

            var changedClassName = "is-active";
            var childClassName = currentNode.getAttribute('child');

            for (var i = 0; i < this.actionNodes.length; i++) {
                if ((currentNode === this.actionNodes[i])) {
                    if (this.oriID !== childClassName) {
                        Office.Controls.Utils.addClass(this.actionNodes[i], changedClassName); 
                        this.oriID = childClassName;
                    } else {
                        this.oriID = "";
                        Office.Controls.Utils.removeClass(this.actionNodes[i], changedClassName); 
                    }
                } else {
                    Office.Controls.Utils.removeClass(this.actionNodes[i], changedClassName);
                }
            }

            for (var i = 0; i < this.actionDetailNodes.length; i++) {
                if (Office.Controls.Utils.containClass(this.actionDetailNodes[i], childClassName) && (this.oriID === childClassName)) {
                    Office.Controls.Utils.addClass(this.actionDetailNodes[i], changedClassName); 
                } else {
                    Office.Controls.Utils.removeClass(this.actionDetailNodes[i], changedClassName);
                } 
            }
        },
    };

    Office.Controls.Persona.PersonaHelper = function () { };
    Office.Controls.Persona.PersonaHelper.createPersona = function (root, templatePath, personaType, dataProvider, isHidden, callback) {
        var personaInstance = new Office.Controls.Persona(root, personaType, dataProvider, isHidden);
        personaInstance.loadTemplateAsync(templatePath, callback);
    };

    Office.Controls.Persona.PersonaHelper.getCachedTemplate = function (templatePath) {
        if ((templatePath === "") || Office.Controls.Utils.isNullOrUndefined(templatePath)) {
            Office.Controls.Utils.errorConsole('Wrong template path');
            return;
        }

        var cacheIndex = templatePath.toLowerCase();
        var cachedTemplate = Office.Controls.Persona.PersonaHelper._cachedTemplates[cacheIndex];
        if (Office.Controls.Utils.isNullOrUndefined(cachedTemplate)) {
            return null;
        } else {
            return cachedTemplate;
        }
    };

    Office.Controls.Persona.PersonaHelper.setCachedTemplate = function (templatePath, xmlDoc) {
        if ((templatePath === "") || Office.Controls.Utils.isNullOrUndefined(templatePath)) {
            Office.Controls.Utils.errorConsole('Wrong template path');
            return;
        }

        if (typeof xmlDoc !== 'object' || (Office.Controls.Utils.isNullOrUndefined(xmlDoc))) {
            Office.Controls.Utils.errorConsole('Invalid template document');
            return;
        }

        var cacheIndex = templatePath.toLowerCase();
        Office.Controls.Persona.PersonaHelper._cachedTemplates[cacheIndex] = xmlDoc;
    };

    Office.Controls.Persona.PersonaHelper.getPersonaType = function () { 
        var personaType = {
            "NameOnly": "nameonly",
            "NameImage" : "nameimage",
            "DetailCard" : "detailcard",
            "PersonaCard" : "personacard"
        };

        return personaType;
    };

    Office.Controls.Persona.ImageSize = function () {};
    Office.Controls.Persona.ImageSize.prototype = {
        s: 0,
        m: 1,
        l: 2
    };

    Office.Controls.Persona.StringUtils = function () { };
    Office.Controls.Persona.StringUtils.getDisplayText = function (displayText, personaType, personaProperty) {
        if (!displayText) {
            return '';
        }
        if (!Office.Controls.Persona.StringUtils._propertyDisplayConfiguration || Office.Controls.Persona.StringUtils._currentPersonaType !== personaType) {
            Office.Controls.Persona.StringUtils._propertyDisplayConfiguration = Office.Controls.Persona.StringUtils._loadLengthConfiguration(personaType);
            Office.Controls.Persona.StringUtils._currentPersonaType = personaType;
        }
        if (Office.Controls.Persona.StringUtils._propertyDisplayConfiguration.length && displayText.length > Office.Controls.Persona.StringUtils._propertyDisplayConfiguration[personaProperty]) {
            return displayText.substr(0, Office.Controls.Persona.StringUtils._propertyDisplayConfiguration[personaProperty]) + '...';
        }
        else {
            return displayText;
        }
    };
    Office.Controls.Persona.StringUtils._loadLengthConfiguration = function (personaType) {
        var returnValue; 

        switch (personaType) {
            case 1:
                returnValue = Office.Controls.Persona.StringUtils._propertyDisplayConfiguration = [ 18, 26, 40, 42 ];
                break;
            case 0:
            case 4:
                if (Office.Controls.Utils.isIE() || Office.Controls.Utils.isFirefox()) {
                    returnValue = Office.Controls.Persona.StringUtils._propertyDisplayConfiguration = [ 30, 0, 40, 42 ];
                }
                else {
                    returnValue = Office.Controls.Persona.StringUtils._propertyDisplayConfiguration = [ 27, 0, 40, 42 ];
                }
                break;
            default:
                returnValue = null;
        }
        return returnValue;
    };

    Office.Controls.Persona.Strings = function() {
    }

    Office.Controls.PersonaResources = function () {
    };

    Office.Controls.PersonaConstants = function () {
    };

    
    if (Office.Controls.Persona.registerClass) { Office.Controls.Persona.registerClass('Office.Controls.Persona'); }
    if (Office.Controls.Persona.Strings.registerClass) { Office.Controls.Persona.Strings.registerClass('Office.Controls.Persona.Strings'); }
    if (Office.Controls.PersonaConstants.registerClass) { Office.Controls.PersonaConstants.registerClass('Office.Controls.PersonaConstants'); }
    if (Office.Controls.PersonaResources.registerClass) { Office.Controls.PersonaResources.registerClass('Office.Controls.PersonaResources'); }
    if (Office.Controls.Persona.PersonaHelper.registerClass) { Office.Controls.Persona.PersonaHelper.registerClass('Office.Controls.Persona.PersonaHelper'); }
    if (Office.Controls.Persona.StringUtils.registerClass) { Office.Controls.Persona.StringUtils.registerClass('Office.Controls.Persona.StringUtils'); }
    Office.Controls.PersonaResources.PersonaName = 'Persona';
    Office.Controls.Persona.Strings.emailString = 'Email';
    Office.Controls.Persona.Strings.lyncMessageString = 'IM';
    Office.Controls.Persona.Strings.phoneString = 'Phone';
    Office.Controls.Persona.Strings.mobileString = 'Mobile';
    Office.Controls.Persona.Strings.workPhoneString = 'Work';
    Office.Controls.Persona.Strings.colonString = ':';
    Office.Controls.Persona.Strings.suspensionPoints = '...';
    Office.Controls.Persona.StringUtils._propertyDisplayConfiguration = null;
    Office.Controls.Persona.StringUtils._currentPersonaType = 0;
    Office.Controls.Persona.PersonaHelper._cachedTemplates = {};
    Office.Controls.PersonaConstants.SectionTag_Main = "persona-section-tag-main";
    Office.Controls.PersonaConstants.SectionTag_Action = "ms-PersonaCard-action";
    Office.Controls.PersonaConstants.SectionTag_ActionDetail = "ms-PersonaCard-actionDetails";
})();