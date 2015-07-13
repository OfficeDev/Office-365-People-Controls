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
            "PrimaryText": 'Cat Miao',
            "SecondaryText": 'Engineer 2, DepartmentA China', // JobTitle, Department
            "TertiaryText": 'BEIJING-Building1-1/XXX', // Office

            "Actions":
                {
                    "Email": "catmiao@companya.com",
                    "WorkPhone": "+86(10) 12345678", 
                    "Mobile" : "+86 1861-0000-000",
                    "Skype" : "catmiao@companya.com",
                }
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
        // Load template & bind data
        this.loadDefaultTemplate(this.templateID);
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
         * Load the given default template
         * @return {[type]}             [description]
         */
        loadDefaultTemplate: function (templateID) {
            var templateNode = Office.Controls.Persona.Templates.DefaultDefinition[templateID].value;
            if (templateNode === "" || (Office.Controls.Utils.isNullOrUndefined(templateNode))) {
                alert('Fail to get the corret template content in loadDefaultTemplate');
                return;
            }
            this.parseTemplate(templateNode);
        },

        /**
         * Parse the persona content loading from template that includes 3 parts:
         *     1. Main: It's a detail card
         *     2. Action bar: It includes the action icons and the click event listener is also attached to each icon.
         *     3. The detail content of each Action icon: When click the icon, the detail shows up.
         * @xmlDoc  {[DomElment} xmlDoc The document loading from template
         */
        parseTemplate: function (templatedContent) {
            try {
                var templateElement = document.createElement("div");
                this.showNode(templateElement, this.isHidden);

                // Get cached view
                var cachedViewWithConstants = Office.Controls.Persona.PersonaHelper.getLocalCache(this.templateID);
                if (cachedViewWithConstants === null)
                {
                    // Replace the constant strings
                    cachedViewWithConstants = this.replaceConstantStrings(templatedContent);
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
         * Show the given domElement node
         * @node  {DomElement}
         * @isHidden  {Boolean}
         */
        showNode: function (node, isHidden) {
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
    Office.Controls.Persona.PersonaHelper.createPersona = function (root, personaObject, personaType) {
        // Make sure the data object is legal.
        var dataObject = Office.Controls.Persona.PersonaHelper.ensurePersonaObjectLegal(personaObject);
        // Create Persona 
        return new Office.Controls.Persona(root, personaType, dataObject, true);
    };

    Office.Controls.Persona.PersonaHelper.createInlinePersona = function (root, personaObject, eventType) {
        var personaCard = null;
        var showNodeQueue = [];
        var personaInstance = Office.Controls.Persona.PersonaHelper.createPersona(root, personaObject, Office.Controls.Persona.PersonaType.TypeEnum.NameImage);
        if (eventType === "click") {
            if (personaInstance.rootNode !== null) {
                Office.Controls.Utils.addEventListener(personaInstance.rootNode, eventType, function (e) {
                    if (personaCard == null) {
                        personaCard = Office.Controls.Persona.PersonaHelper.createPersonaCard(root, personaObject);
                        showNodeQueue.push(personaCard);
                    } else {
                        personaCard.showNode(personaCard.get_rootNode(), (showNodeQueue.length == 0));
                        if (showNodeQueue.length !== 0) {
                            showNodeQueue.pop(personaCard);
                        } else {
                            showNodeQueue.push(personaCard);
                        }
                    }
                });
                Office.Controls.Utils.addEventListener(document, eventType, function () {
                    if (event.target.tagName.toLowerCase() === "html") {
                        if (showNodeQueue.length !== 0) {
                            personaCard.showNode(personaCard.get_rootNode(), false);
                            showNodeQueue.pop(personaCard);
                        }
                    }
                });
            } else {
                Office.Controls.Utils.errorConsole('Wrong template path');
            }
        } 
        return personaInstance;
    };

    Office.Controls.Persona.PersonaHelper.createPersonaCard = function (root, personaObject) {
        return Office.Controls.Persona.PersonaHelper.createPersona(root, personaObject, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard);
    };

    /**
     * Make sure the data object to be used for creating Persona is legal
     */
    Office.Controls.Persona.PersonaHelper.ensurePersonaObjectLegal = function(originalPersonaObj) {
        if (typeof originalPersonaObj !== 'object' || (Office.Controls.Utils.isNullOrUndefined(originalPersonaObj))) {
            Office.Controls.Utils.errorConsole('persona object is null.');
            return;
        }
            
        var personaObj = {};
        personaObj.Id = Office.Controls.Utils.isNullOrUndefined(originalPersonaObj.Id) ? "" : originalPersonaObj.Id;
        personaObj.ImageUrl = Office.Controls.Utils.isNullOrUndefined(originalPersonaObj.ImageUrl) ? Office.Controls.Persona.PersonaHelper._defaultImage : originalPersonaObj.ImageUrl;
        personaObj.PrimaryText = Office.Controls.Utils.isNullOrUndefined(originalPersonaObj.PrimaryText) ? "" : originalPersonaObj.PrimaryText;
        personaObj.SecondaryText = Office.Controls.Utils.isNullOrUndefined(originalPersonaObj.SecondaryText) ? "" : originalPersonaObj.SecondaryText;
        personaObj.SecondaryTextShort = Office.Controls.Persona.StringUtils.getDisplayText(personaObj.SecondaryText, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard, 0);
        personaObj.TertiaryText = Office.Controls.Utils.isNullOrUndefined(originalPersonaObj.TertiaryText) ? "" : originalPersonaObj.TertiaryText;
        personaObj.TertiaryTextShort = Office.Controls.Persona.StringUtils.getDisplayText(personaObj.TertiaryTextShort, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard, 0);

        personaObj.Actions = {};
        personaObj.Actions.Email = Office.Controls.Utils.isNullOrUndefined(originalPersonaObj.Actions.Email) ? "" : originalPersonaObj.Actions.Email;
        personaObj.Actions.WorkPhone = Office.Controls.Utils.isNullOrUndefined(originalPersonaObj.Actions.WorkPhone) ? "" : originalPersonaObj.Actions.WorkPhone;
        personaObj.Actions.Mobile = Office.Controls.Utils.isNullOrUndefined(originalPersonaObj.Actions.Mobile) ? "" : originalPersonaObj.Actions.Mobile;
        personaObj.Actions.Skype = Office.Controls.Utils.isNullOrUndefined(originalPersonaObj.Actions.Skype) ? "" : originalPersonaObj.Actions.Skype;
        
        return personaObj;
    }

    /**
     * Convert AAD User Data To Persona Object
     * @aadUserObject {JSON Object}
     * @return {JSON personaObject}
     */
    Office.Controls.Persona.PersonaHelper.convertAadUserToPersonaObject = function(aadUserObject) {
        if (typeof aadUserObject !== 'object' || (Office.Controls.Utils.isNullOrUndefined(aadUserObject))) {
            Office.Controls.Utils.errorConsole('AAD user data is null.');
            return;
        }

        var displayName = Office.Controls.Persona.StringUtils.getDisplayText(aadUserObject.displayName, Office.Controls.Persona.PersonaType.TypeEnum.NameImage, 3);
            
        var personaObj = {};
        personaObj.Id = aadUserObject.personId;
        personaObj.ImageUrl = Office.Controls.Persona.PersonaHelper._defaultImage;
        personaObj.PrimaryText = (displayName === "") ? Office.Controls.Persona.Strings.EmptyDisplayName : displayName;
        
        if (aadUserObject.jobTitle !== null) {
            personaObj.SecondaryText = aadUserObject.jobTitle  + Office.Controls.Persona.Strings.Comma + aadUserObject.department;
        } else {
            personaObj.SecondaryText = aadUserObject.department;   
        }
        personaObj.SecondaryTextShort = Office.Controls.Persona.StringUtils.getDisplayText(personaObj.SecondaryText, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard, 0);
        personaObj.TertiaryText = Office.Controls.Persona.StringUtils.getDisplayText(aadUserObject.office, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard, 0);

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
                    returnValue = Office.Controls.Persona.StringUtils._propertyDisplayConfiguration = [ 30, 0, 40, 42 ];
                }
                break;
            default:
                returnValue = null;
        }
        return returnValue;
    };

    Office.Controls.Persona.Strings = function() {
    }

    Office.Controls.Persona.Templates = function() {
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
    Office.Controls.Persona.PersonaHelper._defaultImage = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAABGdBTUEAALGPC/xhBQAAAAFzUkdCAK7OHOkAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAORQTFRFsbGxv7+/wcHBu7u72dnZ7u7u/////Pz85ubmy8vLtbW1wsLC5OTk/f3939/ft7e34eHh+vr61NTUs7OzwMDA8/Pz7OzsxMTE9PT0vLy8+/v7+fn5urq6vb298fHxtLS0srKyzc3N9/f32NjYtra209PT7e3t7+/v/v7+yMjI3t7e4+Pj6+vr6urq4uLi19fXx8fH0tLS9fX1ubm59vb2z8/PuLi4+Pj429vbzs7O0NDQ4ODg1dXV2tra1tbWzMzMxcXF6enp0dHR8vLyysrKxsbG3d3d6Ojovr6+5eXl5+fn8PDwYCkYCwAAAAFiS0dEBmFmuH0AAAAJcEhZcwAAdTAAAHUwAd0zcs0AAAKgSURBVGje7ZjdVtpQEIUByxCIIRKtYECgSSVULbWopRDBarVq+/7vU1raBcTMmbEzueha7Ous/XEOZ35zuY022mij/0z5QqGQmfnWqyL8VskqV/Tt7W1YkVN1de0rO5BQzdtV9N97Dc+1r3eIegnS1DhQ8rd9SFezpeJ/2ARM7Y6Cf7cIuCwFgAcmvRH7t3wjIBCHXBXMCoX+rkMA/LcywBFQOpIBeiQgkt1QnwSAKNre0f6ylxoyAMcSwAkDsC8BnDIAgQQQMQCOBNBmAECSUotZA95zAJIrGjD8mxLABwagKAGcMQADCYCTiz5KAIxsCnkRwCP9hyJ/uqKJayb1UJ1zIeDc3FXAjtCfCoWa9ADzzv3CBLgU+8/HphLuf6rgn8uVUf/GJxUA2hwFOu37rzOk3tJQawCZy075p7eFTeO6OmHiECN5455Q63OwtG+PdWbM+qQaXx0uL2oa96IoGsxWE+h1bB1P7H9x3wpHi59bNQwZ7p+q3fZeOtR+Wan3DXQ/YS+vzYlf8mjd9Z7RuUr9ajdcS+ZOzE5Ml7Xkm7xJidn8s7bMP2PZd+OUqGom01onTKlF/VsOAJn7vpZXnmZrFqR/dUf7j9Hc5lu39UIlfz+Jh/g3ZA/g4psDlhrU5EwNxqSILuCA0WqZRUzOU6k/gPEldUdyQM8E4HS7lPqmtMSZW0mZSsWNBsA0OQdye+P6oqLhDz4O2FMBAJ6273QAdRRAzxssTVDANx3AFAVYOgB88OQsPxh6QAHCWvBX6GyuEwbzooMBHpUAaKTdKwEA61zHcuuFsMlnpgV4zDbO8FzxpAXA2i/Wio4jbHy+kFsv5CEAcrXCFdJ8uVr+WFXOqwFOMg5k+J5xIGMbebVAhh8ZBzK2KVQLZKituP4EqRB824c6sq4AAAAldEVYdGRhdGU6Y3JlYXRlADIwMTQtMTAtMjlUMjA6MjQ6MTktMDU6MDBCpOLkAAAAJXRFWHRkYXRlOm1vZGlmeQAyMDE0LTEwLTI5VDIwOjI0OjE5LTA1OjAwM/laWAAAAABJRU5ErkJggg==";
    Office.Controls.PersonaConstants.SectionTag_Main = "persona-section-tag-main";
    Office.Controls.PersonaConstants.SectionTag_Action = "ms-PersonaCard-action";
    Office.Controls.PersonaConstants.SectionTag_ActionDetail = "ms-PersonaCard-actionDetails";
    Office.Controls.Persona.Templates.DefaultDefinition = {
        "nameonly": 
        {
            value: "<div class=\"ms-Persona ms-Persona--tiny readOnly clickStyle\" AriaTabIndex=\"0\"><div class=\"ms-Persona-primaryText nameOnlyText\"><Label Text=\"${PrimaryText}\">${PrimaryText}</Label></div></div>",
        },
        "nameimage": 
        {
            value: "<div class=\"ms-Persona\"><div class=\"image\"><image class=\"imageOfNameImage\" ImageName=\"${PrimaryText}\" Src=\"${ImageUrl}\"></image></div><div class=\"ms-Persona-details ms-Persona-details-nameImage\"><div class=\"ms-Persona-primaryText ms-Persona-primaryText-nameImage\"><Label class=\"clickStyle\" Text=\"${PrimaryText}\" Title=\"${PrimaryText}\">${PrimaryText}</Label></div><div class=\"ms-Persona-secondaryText ms-Persona-secondaryText-nameImage\"><Label class=\"defaultStyle\" Text=\"${SecondaryText}\" Title=\"${SecondaryText}\">${SecondaryText}</Label></div></div></div>",
        },
        "detailcard": 
        {
            value: "<div class=\"ms-PersonaCard personaCard-customized detail displayMode\"><div class=\"ms-PersonaCard-persona\"><div class=\"ms-Persona ms-Persona--xl\"><image class=\"ms-Persona-image image\" ImageName=\"${PrimaryText}\" Src=\"${ImageUrl}\"></image><div class=\"ms-Persona-details\"><div class=\"ms-Persona-primaryText\"><Label class=\"defaultStyle\" Text=\"${PrimaryText}\">${PrimaryText}</Label></div><div class=\"ms-Persona-secondaryText\"><Label class=\"defaultStyle\" Text=\"${SecondaryText}\">${SecondaryText}</Label></div><div class=\"ms-Persona-tertiaryText\"><Label class=\"defaultStyle\" Text=\"${TertiaryText}\">${TertiaryText}</Label></div></div></div></div></div>",
        },
        "personacard": 
        {
            value: "<div class=\"ms-PersonaCard personaCard-customized detail displayMode\"><div class=\"ms-PersonaCard-persona persona-section-tag-main\"><div class=\"ms-Persona ms-Persona--xl\"><image class=\"ms-Persona-image image\" ImageName=\"${PrimaryText}\" Src=\"${ImageUrl}\"></image><div class=\"ms-Persona-details\"><div class=\"ms-Persona-primaryText\"><Label class=\"defaultStyle\" Text=\"${PrimaryText}\">${PrimaryText}</Label></div><div class=\"ms-Persona-secondaryText\"><Label class=\"defaultStyle\" title=\"${SecondaryText}\">${SecondaryTextShort}</Label></div><div class=\"ms-Persona-tertiaryText\"><Label class=\"defaultStyle\" Text=\"${TertiaryText}\">${TertiaryText}</Label></div></div></div></div><ul class=\"ms-PersonaCard-actions\"><li class=\"ms-PersonaCard-action\" child=\"action-detail-mail\"><i class=\"ms-Icon ms-Icon--mail icon\"><span></span></i></li><li class=\"ms-PersonaCard-action\" child=\"action-detail-phone\"><i class=\"ms-Icon ms-Icon--phone icon\"><span></span></i></li><li class=\"ms-PersonaCard-action\" child=\"action-detail-chat\"><i class=\"ms-Icon ms-Icon--chat icon\"><span></span></i></li></ul><div class=\"ms-PersonaCard-actionDetails action-detail-mail\"><div class=\"ms-PersonaCard-detailLine\"><span class=\"ms-PersonaCard-detailLabel\">${Strings.Label.Email}</span><a href=\"${Strings.Protocol.Email}${Actions.Email}\">${Actions.Email}</a></div></div><div class=\"ms-PersonaCard-actionDetails action-detail-phone\"><div class=\"ms-PersonaCard-detailLine\"><span class=\"ms-PersonaCard-detailLabel\">${Strings.Label.WorkPhone}</span><a href=\"${Strings.Protocol.Phone}${Actions.WorkPhone}\">${Actions.WorkPhone}</a><br/><span class=\"ms-PersonaCard-detailLabel\">${Strings.Label.Mobile}</span><a href=\"${Strings.Protocol.Phone}${Actions.Mobile}\">${Actions.Mobile}</a></div></div><div class=\"ms-PersonaCard-actionDetails action-detail-chat\"><div class=\"ms-PersonaCard-detailLine\"><span class=\"ms-PersonaCard-detailLabel\">${Strings.Label.Skype}</span><a href=\"${Strings.Protocol.Skype}${Actions.Skype}\">${Actions.Skype}</a></div></div></div>",
        },
    };
})();