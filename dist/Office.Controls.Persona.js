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
    *  The format of personaObject:
    * {
            "Id": "[guid]",
            "ImageUrl": "control/images/icon.png",
            "PrimaryText": '***REMOVED*** Chen',
            "SecondaryText": 'Software Engineer 2, ASG EA China', // JobTitle, Department
            "TertiaryText": 'BEIJING-BJW-1/12329', // Office

            "Actions":
            {
                "Email": "***REMOVED***@microsoft.com",
                "WorkPhone": "+86(10) 59173216", 
                "Mobile" : "+86 1861-2947-014",
                "Skype" : "***REMOVED***@microsoft.com",
            }
        };
    */
    Office.Controls.Persona = function (root, personaType, personaObject, isHidden) {
        if (typeof root !== 'object' || typeof personaType !== 'string' || typeof personaObject !== 'object') {
                Office.Controls.Utils.errorConsole('Invalid parameters type');
                return;
        }
            
        this.root = root;
        this.templateID = personaType.toString();
        this.personaObject = personaObject;
        this.isHidden = isHidden;
    };

    Office.Controls.Persona.prototype = {
        onError: null,
        rootNode: null,
        mainNode: null,
        actionNodes: null,
        actionDetailNodes: null,
        constantObject: {}, 
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
         * @callback  {Function} callback(errorMessage, elementNode)
         * @return {null}
         */
        loadTemplateAsync: function (templatePath, callback) {
            var self = this;
            var cachedTemplate = Office.Controls.Persona.PersonaHelper.getLocalCache(templatePath);

            // Get cache
            if (cachedTemplate !== null)
            {
               self.parseTemplate(cachedTemplate);
               callback(self.rootNode, null);
               return;
            } 

            var xhr = new XMLHttpRequest();
            xhr.open("GET", templatePath, true);
            xhr.onreadystatechange = function() {
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
                        Office.Controls.Persona.PersonaHelper.setLocalCache(templatePath, xmlDoc);
                        self.parseTemplate(xmlDoc);
                        callback(self.rootNode, null); 
                    } else {
                        callback(null, "Unknown error");
                        return; 
                    }
                }
            };
            xhr.send();
        },

        /**
         * Parse the persona content loading from template that includes 3 parts:
         *     1. Main: It's a detail card
         *     2. Action bar: It includes the action icons and the click event listener is also attached to each icon.
         *     3. The detail content of each Action icon: When click the icon, the detail shows up.
         * @xmlDoc  {[DomElment} xmlDoc The document loading from template
         */
        parseTemplate: function (xmlDoc) {
            try {
                var templateNode = xmlDoc.getElementById(this.templateID);
                var templateElement = document.createElement("div");
                this.hiddenNode(templateElement, this.isHidden);

                // Get cached view
                var cachedViewWithConstants = Office.Controls.Persona.PersonaHelper.getLocalCache(this.templateID);
                if (cachedViewWithConstants === null)
                {
                    // Replace the constant strings
                    cachedViewWithConstants = this.replaceConstantStrings(templateNode.innerHTML);
                    // Save view to local cache
                    Office.Controls.Persona.PersonaHelper.setLocalCache(this.templateID, cachedViewWithConstants);
                }

                // Bind the business data
                templateElement.innerHTML = this.bindPersonaData(cachedViewWithConstants);
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
         * Bind businees data to template
         * @htmlStr  {string}
         * @return {string}
         */
        bindPersonaData: function (htmlStr) {
            var regExp = /\$\{([^\}\{]+)\}/g;
            return this.bindData(htmlStr, regExp, this.personaObject);
        },

        /**
         * Replace constant strings to template
         * @htmlStr  {string}
         * @return {string}
         */
        replaceConstantStrings : function (htmlStr) {
            // Init constant strings
            this.initiateStringObject();

            var regExp = /\$\{Strings([^\}\{]+)\}/g;
            return this.bindData(htmlStr, regExp, this.constantObject);
        },

         /**
          * Bind generic data to template
          * @htmlStr  {string}
          * @regExp  {string}
          * @dataObject  {JsonObject}
          * @return {string}
          */
        bindData : function(htmlStr, regExp, dataObject)
        {
            if ((htmlStr === "") || Office.Controls.Utils.isNullOrUndefined(htmlStr) || (regExp === "") || Office.Controls.Utils.isNullOrUndefined(regExp) || (typeof dataObject !== 'object') || Office.Controls.Utils.isNullOrUndefined(dataObject)) {
                Office.Controls.Utils.errorConsole('data object is null.');
                return htmlStr;
            }
            
            var resultStr = htmlStr;
            var propertyName, propertyValue;
            var self = this;

            // Get the property names
            var properties = resultStr.match(regExp);
            if (properties !== null) {
                for (var i = 0; i < properties.length; i++) { 
                    propertyName = properties[i].substring(2, properties[i].length - 1);
                    propertyValue = self.getValueFromJSONObject(dataObject, propertyName);
                    resultStr = resultStr.replace(properties[i], propertyValue);
                }
            }

            return resultStr;
        },

        /**
         * Strings:
         * {
            "Label":{
                        "Email": "Work: "
                        "WorkPhone": "Work: ",
                        "Mobile": "Mobile: ",
                        "Skype": "Skype: "
                    },

            "Protocol": {
                            "Email": "mailto:",
                            "Phone": "tel:",
                            "Skype": "sip:",
                        }
            }
         */
        initiateStringObject : function()
        {
            var colonSpace = Office.Controls.Persona.Strings.Colon + Office.Controls.Persona.Strings.Space;

            this.constantObject.Strings = {};
            this.constantObject.Strings.Label = {};
            this.constantObject.Strings.Label.Email = Office.Controls.Persona.Strings.Label_Work + colonSpace;
            this.constantObject.Strings.Label.WorkPhone = Office.Controls.Persona.Strings.Label_Work + colonSpace;
            this.constantObject.Strings.Label.Mobile = Office.Controls.Persona.Strings.Label_Mobile + colonSpace
            this.constantObject.Strings.Label.Skype = Office.Controls.Persona.Strings.Label_Skype + colonSpace;
            
            this.constantObject.Strings.Protocol = {};
            this.constantObject.Strings.Protocol.Email = Office.Controls.Persona.Strings.Protocol_Mail;
            this.constantObject.Strings.Protocol.Phone = Office.Controls.Persona.Strings.Protocol_Phone;
            this.constantObject.Strings.Protocol.Skype = Office.Controls.Persona.Strings.Protocol_Skype;
        },

        /**
         * Parse the json object to get the corresponding value
         * @objectName  {string}
         * @return {object}
         */
        getValueFromJSONObject: function (obj, objectName) {
            var rtnValue =  Office.Controls.Utils.getObjectFromJSONObjectName(obj, objectName);
            if (rtnValue === null) {
                Office.Controls.Utils.errorConsole('the json object is null for data binding.');
                return;
            } 

            return rtnValue;
        },

        /**
         * Hidden the given domElement node
         * @node  {DomElement}
         * @isHidden  {Boolean}
         */
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

        /**
         * [removeDetailCard description]
         * @return {[type]}
         */
        removeDetailCard: function () {
        },

        /**
         * Set the style of ative action button
         * @event {HtmlEvent}
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
    Office.Controls.Persona.PersonaHelper.createPersona = function (root, templatePath, personaType, personaObject, isHidden, callback) {
        var personaInstance = new Office.Controls.Persona(root, personaType, personaObject, isHidden);
        personaInstance.loadTemplateAsync(templatePath, callback);
    };

    Office.Controls.Persona.PersonaHelper.createInlinePersona = function (root, templatePath, personaObject, eventType) {
        var personaCard = null;
        var isShow = true;

        var personaInstance = new Office.Controls.Persona(root, Office.Controls.Persona.PersonaType.TypeEnum.NameImage, personaObject, true);
        personaInstance.loadTemplateAsync(templatePath, function callback(rootNode, error) {
            if (eventType === "click") {
                if (rootNode !== null) {
                    Office.Controls.Utils.addEventListener(rootNode, eventType, function (e) {
                        if (personaCard == null) {
                            personaCard = new Office.Controls.Persona(root, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard, personaObject, true);
                            personaCard.loadTemplateAsync(tempPath, function (rootNode, error) {
                                isShow = false;
                            }); 
                        } else {
                            personaCard.hiddenNode(personaCard.get_rootNode(), isShow);
                            isShow = isShow ? false : true;
                        }
                    });
                } else {
                    Office.Controls.Utils.errorConsole('Wrong template path');
                }
            } 
        });
        return personaInstance;
    };

    Office.Controls.Persona.PersonaHelper.createPersonaCard = function (root, templatePath, personaObject, callback) {
        return Office.Controls.Persona.PersonaHelper.createPersona(root, templatePath, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard, personaObject, true, callback);
    };

    /**
     * Convert AAD User Data To Persona Object
     * @aadUserObject {JSON Object}
     * @return {JSON Object}
     * {
     *  "Id": "[guid]",
        "ImageUrl": "control/images/icon.png",
        "PrimaryText": '***REMOVED*** Chen',
        "SecondaryText": 'Software Engineer 2, ASG EA China', // JobTitle, Department
        "TertiaryText": 'BEIJING-BJW-1/12329', // Office

        "Actions":
            {
                "Email": "***REMOVED***@microsoft.com",
                "WorkPhone": "+86(10) 59173216", 
                "Mobile" : "+86 1861-2947-014",
                "Skype" : "***REMOVED***@microsoft.com",
            }
        }
     */
    Office.Controls.Persona.PersonaHelper.convertAadUserToPersonaObject = function(aadUserObject) {
        if (typeof aadUserObject !== 'object' || (Office.Controls.Utils.isNullOrUndefined(aadUserObject))) {
            Office.Controls.Utils.errorConsole('AAD user data is null.');
            return;
        }

        var displayName = Office.Controls.Persona.StringUtils.getDisplayText(aadUserObject.displayName, Office.Controls.Persona.PersonaType.TypeEnum.NameImage, 3);
            
        var personaObj = {};
        personaObj.Id = aadUserObject.personId;
        personaObj.ImageUrl = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAABGdBTUEAALGPC/xhBQAAAAFzUkdCAK7OHOkAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAORQTFRFsbGxv7+/wcHBu7u72dnZ7u7u/////Pz85ubmy8vLtbW1wsLC5OTk/f3939/ft7e34eHh+vr61NTUs7OzwMDA8/Pz7OzsxMTE9PT0vLy8+/v7+fn5urq6vb298fHxtLS0srKyzc3N9/f32NjYtra209PT7e3t7+/v/v7+yMjI3t7e4+Pj6+vr6urq4uLi19fXx8fH0tLS9fX1ubm59vb2z8/PuLi4+Pj429vbzs7O0NDQ4ODg1dXV2tra1tbWzMzMxcXF6enp0dHR8vLyysrKxsbG3d3d6Ojovr6+5eXl5+fn8PDwYCkYCwAAAAFiS0dEBmFmuH0AAAAJcEhZcwAAdTAAAHUwAd0zcs0AAAKgSURBVGje7ZjdVtpQEIUByxCIIRKtYECgSSVULbWopRDBarVq+/7vU1raBcTMmbEzueha7Ous/XEOZ35zuY022mij/0z5QqGQmfnWqyL8VskqV/Tt7W1YkVN1de0rO5BQzdtV9N97Dc+1r3eIegnS1DhQ8rd9SFezpeJ/2ARM7Y6Cf7cIuCwFgAcmvRH7t3wjIBCHXBXMCoX+rkMA/LcywBFQOpIBeiQgkt1QnwSAKNre0f6ylxoyAMcSwAkDsC8BnDIAgQQQMQCOBNBmAECSUotZA95zAJIrGjD8mxLABwagKAGcMQADCYCTiz5KAIxsCnkRwCP9hyJ/uqKJayb1UJ1zIeDc3FXAjtCfCoWa9ADzzv3CBLgU+8/HphLuf6rgn8uVUf/GJxUA2hwFOu37rzOk3tJQawCZy075p7eFTeO6OmHiECN5455Q63OwtG+PdWbM+qQaXx0uL2oa96IoGsxWE+h1bB1P7H9x3wpHi59bNQwZ7p+q3fZeOtR+Wan3DXQ/YS+vzYlf8mjd9Z7RuUr9ajdcS+ZOzE5Ml7Xkm7xJidn8s7bMP2PZd+OUqGom01onTKlF/VsOAJn7vpZXnmZrFqR/dUf7j9Hc5lu39UIlfz+Jh/g3ZA/g4psDlhrU5EwNxqSILuCA0WqZRUzOU6k/gPEldUdyQM8E4HS7lPqmtMSZW0mZSsWNBsA0OQdye+P6oqLhDz4O2FMBAJ6273QAdRRAzxssTVDANx3AFAVYOgB88OQsPxh6QAHCWvBX6GyuEwbzooMBHpUAaKTdKwEA61zHcuuFsMlnpgV4zDbO8FzxpAXA2i/Wio4jbHy+kFsv5CEAcrXCFdJ8uVr+WFXOqwFOMg5k+J5xIGMbebVAhh8ZBzK2KVQLZKituP4EqRB824c6sq4AAAAldEVYdGRhdGU6Y3JlYXRlADIwMTQtMTAtMjlUMjA6MjQ6MTktMDU6MDBCpOLkAAAAJXRFWHRkYXRlOm1vZGlmeQAyMDE0LTEwLTI5VDIwOjI0OjE5LTA1OjAwM/laWAAAAABJRU5ErkJggg==";
        personaObj.PrimaryText = (displayName === "") ? Office.Controls.Persona.Strings.EmptyDisplayName : displayName;
        if (aadUserObject.jobTitle !== null) {
            personaObj.SecondaryText = aadUserObject.jobTitle  + Office.Controls.Persona.Strings.Comma + aadUserObject.department;
        } else {
            personaObj.SecondaryText = aadUserObject.department;   
        }
        personaObj.SecondaryText = Office.Controls.Persona.StringUtils.getDisplayText(personaObj.SecondaryText, Office.Controls.Persona.PersonaType.TypeEnum.NameImage, 3);
        personaObj.TertiaryText = Office.Controls.Persona.StringUtils.getDisplayText(aadUserObject.office, Office.Controls.Persona.PersonaType.TypeEnum.NameImage, 3);

        personaObj.Actions = {};
        personaObj.Actions.Email = aadUserObject.mail;
        personaObj.Actions.WorkPhone = aadUserObject.workPhone;
        personaObj.Actions.Mobile = aadUserObject.mobile;
        personaObj.Actions.Skype = aadUserObject.sipAddress;
        
        return personaObj;
    }

    Office.Controls.Persona.PersonaHelper.convertAadUsersToPersonaObjects = function(aadUsers) {
        if (typeof aadUsers !== 'object' || (Office.Controls.Utils.isNullOrUndefined(aadUsers))) {
            Office.Controls.Utils.errorConsole('AAD user collection is null.');
            return;
        }

        var personaObjects = [];
        aadUsers.forEach(function (aadUser) {
            personaObjects.push(Office.Controls.Persona.PersonaHelper.convertAadUserToPersonaObject(aadUser));
        });
        return personaObjects;
    }

    Office.Controls.Persona.PersonaHelper.getLocalCache = function (cacheId) {
        if ((cacheId === "") || Office.Controls.Utils.isNullOrUndefined(cacheId)) {
            Office.Controls.Utils.errorConsole('Wrong Cache ID');
            return;
        }

        var cacheIndex = cacheId.toLowerCase();
        var cachedObject = Office.Controls.Persona.PersonaHelper._localCache[cacheIndex];
        if (Office.Controls.Utils.isNullOrUndefined(cachedObject)) {
            return null;
        } else {
            return cachedObject;
        }
    };

    Office.Controls.Persona.PersonaHelper.setLocalCache = function (cacheId, object) {
        if ((cacheId === "") || Office.Controls.Utils.isNullOrUndefined(cacheId)) {
            Office.Controls.Utils.errorConsole('Wrong Cache ID');
            return;
        }

        if ((typeof object !== 'object' && typeof object !== 'string') || (Office.Controls.Utils.isNullOrUndefined(object))) {
            Office.Controls.Utils.errorConsole('Invalid Cached Object');
            return;
        }

        var cacheIndex = cacheId.toLowerCase();
        Office.Controls.Persona.PersonaHelper._localCache[cacheIndex] = object;
    };

    // The Persona Type
    Office.Controls.Persona.PersonaType = function() {};
    Office.Controls.Persona.PersonaType.TypeEnum = {
        NameOnly: "nameonly",
        NameImage: "nameimage",
        DetailCard: "detailcard",
        PersonaCard: "personacard",
    };

    // The Persona Type
    Office.Controls.Persona.ImageSize = function() {};
    Office.Controls.Persona.ImageSize.TypeEnum = {
        s: 0,
        m: 1,
        l: 2,
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
            case Office.Controls.Persona.PersonaType.TypeEnum.NameImage:
                returnValue = Office.Controls.Persona.StringUtils._propertyDisplayConfiguration = [ 18, 26, 40, 42 ];
                break;
            case Office.Controls.Persona.PersonaType.TypeEnum.NameOnly:
            case Office.Controls.Persona.PersonaType.TypeEnum.DetailCard:
            case Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard:
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
    Office.Controls.Persona.Strings.Label_Skype = 'Skype';
    Office.Controls.Persona.Strings.Label_Work = 'Work';
    Office.Controls.Persona.Strings.Label_Mobile = 'Mobile';
    Office.Controls.Persona.Strings.Protocol_Mail = 'mailto:';
    Office.Controls.Persona.Strings.Protocol_Phone = 'tel:';
    Office.Controls.Persona.Strings.Protocol_Skype = 'sip:';
    Office.Controls.Persona.Strings.Colon = ':';
    Office.Controls.Persona.Strings.Comma = ',';
    Office.Controls.Persona.Strings.Space = ' ';
    Office.Controls.Persona.Strings.SuspensionPoints = '...';
    Office.Controls.Persona.Strings.EmptyDisplayName = 'No Name';
    Office.Controls.Persona.StringUtils._propertyDisplayConfiguration = null;
    Office.Controls.Persona.StringUtils._currentPersonaType = 0;
    Office.Controls.Persona.PersonaHelper._localCache = {};
    Office.Controls.PersonaConstants.SectionTag_Main = "persona-section-tag-main";
    Office.Controls.PersonaConstants.SectionTag_Action = "ms-PersonaCard-action";
    Office.Controls.PersonaConstants.SectionTag_ActionDetail = "ms-PersonaCard-actionDetails";
})();