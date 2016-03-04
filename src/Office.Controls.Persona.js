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
            "id": "person id",
            "imgSrc": "",
            "primaryText": 'Jerry Anderson',
            "secondaryText": 'Software Engineer, DepartmentA China', // JobTitle, Department
            "tertiaryText": 'BEIJING-Building1-1/XXX', // Office

            "actions":
                {
                    "email": "jerrya@companya.com",
                    "workPhone": "+86(10) 12345678", 
                    "mobile" : "+86 1861-0000-000",
                    "skype" : "jerrya@companya.com",
                }
            }
        };
    */
    Office.Controls.Persona = function (root, personaType, personaObject, isHidden, res) {
        if (typeof root !== 'object' || typeof personaType !== 'string' || typeof personaObject !== 'object') {
                Office.Controls.Utils.errorConsole('Invalid parameters type');
                return;
        }
            
        this.root = root;
        this.templateID = personaType.toString();
        this.personaObject = personaObject;
        this.isHidden = isHidden;
        
        if (!Office.Controls.Utils.isNullOrUndefined(res)) {
            Office.Controls.Persona.PersonaHelper._resourceStrings = res;
        } 
        
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
                // Replace the constant strings 
                // that can't use cache in the case of creating multiple ones
                cachedViewWithConstants = this.replaceConstantStrings(templatedContent);
                
                if (cachedViewWithConstants === null)
                {
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

            var regExp = /\$\{strings([^\}\{]+)\}/g;
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
         * strings:
         * {
            "label":{
                        "email": "Work: "
                        "workPhone": "Work: ",
                        "mobile": "Mobile: ",
                        "skype": "Skype: "
                    },

            "protocol": {
                            "email": "mailto:",
                            "phone": "tel:",
                            "skype": "sip:",
                        }
            }
         */
        initiateStringObject : function() {
            this.constantObject.strings = {};
            this.constantObject.strings.label = {};
            this.constantObject.strings.label.email = this.getResourceString("Email");
            this.constantObject.strings.label.workPhone = this.getResourceString("WorkPhone");
            this.constantObject.strings.label.mobile = this.getResourceString("Mobile");
            this.constantObject.strings.label.skype = this.getResourceString("Skype");
            
            this.constantObject.strings.protocol = {};
            this.constantObject.strings.protocol.email = Office.Controls.Persona.Strings.Protocol_Mail;
            this.constantObject.strings.protocol.phone = Office.Controls.Persona.Strings.Protocol_Phone;
            this.constantObject.strings.protocol.skype = Office.Controls.Persona.Strings.Protocol_Skype;
        },
        
        getResourceString : function(resName, res) {
            return Office.Controls.Persona.PersonaHelper.getResourceString(resName, res);
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
        }
    };

    Office.Controls.Persona.PersonaHelper = function () { };
    /**
     * [createPersona description]
     * @param  {[type]} root         [description]
     * @param  {[type]} personObject same as peoplepicker's personObject queried from AAD
     * @param  {[type]} personaType  [description]
     * @return {[type]}              [description]
     */
    Office.Controls.Persona.PersonaHelper.createPersona = function (root, personObj, personaType, res) {
        // Make sure the data object is legal.
        var personaObj = Office.Controls.Persona.PersonaHelper.convertAadUserToPersonaObject(personObj, res);
        var dataObj = Office.Controls.Persona.PersonaHelper.ensurePersonaObjectLegal(personaObj, personaType);
        // Create Persona 
        return new Office.Controls.Persona(root, personaType, dataObj, true, res);
    };

    Office.Controls.Persona.PersonaHelper.createInlinePersona = function (root, personObject, eventType, res) {
        var personaCard = null;
        var showNodeQueue = Office.Controls.Persona.PersonaHelper._showNodeQueue;
        var personaInstance = Office.Controls.Persona.PersonaHelper.createPersona(root, personObject, Office.Controls.Persona.PersonaType.TypeEnum.NameImage, res);
        if (eventType === "click") {
            if (personaInstance.rootNode !== null) {
                Office.Controls.Utils.addEventListener(personaInstance.rootNode, eventType, function (e) {
                    if (personaCard == null) {
                        // Close other instances on the same page and keep one instance show at most
                        if (showNodeQueue.length !== 0) {
                            var nodeItem = showNodeQueue.pop();
                            nodeItem.showNode(nodeItem.get_rootNode(), false);
                        }
                        personaCard = Office.Controls.Persona.PersonaHelper.createPersonaCard(root, personObject, res);
                        showNodeQueue.push(personaCard);
                    } else {
                        if (showNodeQueue.length !== 0) {
                            var nodeItem = showNodeQueue.pop();
                            nodeItem.showNode(nodeItem.get_rootNode(), false);
                            if (nodeItem !== personaCard) {
                                personaCard.showNode(personaCard.get_rootNode(),true);
                                showNodeQueue.push(personaCard);
                            } 
                        } else {
                            personaCard.showNode(personaCard.get_rootNode(),true);        
                            showNodeQueue.push(personaCard);
                        }
                    }
                });
                Office.Controls.Utils.addEventListener(document, eventType, function () {
                    if (event.target.tagName.toLowerCase() === "html") {
                        if (showNodeQueue.length !== 0) {
                            var nodeItem = showNodeQueue.pop();
                            nodeItem.showNode(nodeItem.get_rootNode(), false);
                        }
                    }
                });
            } else {
                Office.Controls.Utils.errorConsole('Wrong template path');
            }
        } 
        return personaInstance;
    };

    Office.Controls.Persona.PersonaHelper.createImageOnlyPersona = function (root, personObject, eventType, res, dataLoader) {
        var personaCard = null;
        var showNodeQueue = Office.Controls.Persona.PersonaHelper._showNodeQueue;
        var personaInstance = Office.Controls.Persona.PersonaHelper.createPersona(root, personObject, Office.Controls.Persona.PersonaType.TypeEnum.ImageOnly, res);
        if (eventType === "click") {
            if (personaInstance.rootNode !== null) {
                Office.Controls.Utils.addEventListener(personaInstance.rootNode, eventType, function (e) {
                    if (personaCard == null) {
                        // Close other instances on the same page and keep one instance show at most
                        if (showNodeQueue.length !== 0) {
                            var nodeItem = showNodeQueue.pop();
                            nodeItem.showNode(nodeItem.get_rootNode(), false);
                        }
                        // If the data loader function defined, need to load the full data before rendering
                        if (dataLoader != null) {
                            dataLoader(personObject, function (personObjectFull) {
                                personaCard = Office.Controls.Persona.PersonaHelper.createPersonaCard(root, personObjectFull, res);
                                showNodeQueue.push(personaCard);
                            });
                        }
                        else {
                            personaCard = Office.Controls.Persona.PersonaHelper.createPersonaCard(root, personObject, res);
                            showNodeQueue.push(personaCard);
                        }
                        
                    } else {
                        if (showNodeQueue.length !== 0) {
                            var nodeItem = showNodeQueue.pop();
                            nodeItem.showNode(nodeItem.get_rootNode(), false);
                            if (nodeItem !== personaCard) {
                                personaCard.showNode(personaCard.get_rootNode(), true);
                                showNodeQueue.push(personaCard);
                            }
                        } else {
                            personaCard.showNode(personaCard.get_rootNode(), true);
                            showNodeQueue.push(personaCard);
                        }
                    }
                });
                Office.Controls.Utils.addEventListener(document, eventType, function () {
                    if (event.target.tagName.toLowerCase() === "html") {
                        if (showNodeQueue.length !== 0) {
                            var nodeItem = showNodeQueue.pop();
                            nodeItem.showNode(nodeItem.get_rootNode(), false);
                        }
                    }
                });
            } else {
                Office.Controls.Utils.errorConsole('Wrong template path');
            }
        }
        return personaInstance;
    };

    Office.Controls.Persona.PersonaHelper.createPersonaCard = function (root, personObject, res) {
        return Office.Controls.Persona.PersonaHelper.createPersona(root, personObject, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard, res);
    };

    /**
     * Make sure the data object to be used for creating Persona is legal
     */
    Office.Controls.Persona.PersonaHelper.ensurePersonaObjectLegal = function(oriPersonaObj, personaType) {
        // Get the definition of standard object
        var personaObj = Office.Controls.Persona.PersonaHelper.ensureJsonObjectLegal(Office.Controls.Persona.PersonaHelper.getPersonaObjectDef(), oriPersonaObj);
        personaObj.imgSrc = Office.Controls.Persona.StringUtils.setNullOrUndefinedAsEmpty(oriPersonaObj.imgSrc, Office.Controls.Persona.PersonaHelper._defaultImage);
        
        personaObj.primaryTextShort = Office.Controls.Persona.StringUtils.getDisplayText(personaObj.primaryText, personaType, 0);
        personaObj.secondaryTextShort = Office.Controls.Persona.StringUtils.getDisplayText(personaObj.secondaryText, personaType, 2);
        personaObj.tertiaryTextShort = Office.Controls.Persona.StringUtils.getDisplayText(personaObj.tertiaryText, personaType, 2);
        
        personaObj.actions.emailShort = Office.Controls.Persona.StringUtils.getDisplayText(personaObj.actions.email, personaType, 3);
        personaObj.actions.workPhoneShort = Office.Controls.Persona.StringUtils.getDisplayText(personaObj.actions.workPhone, personaType, 3);
        personaObj.actions.mobileShort = Office.Controls.Persona.StringUtils.getDisplayText(personaObj.actions.mobile, personaType, 3);
        personaObj.actions.skypeShort = Office.Controls.Persona.StringUtils.getDisplayText(personaObj.actions.skype, personaType, 3);
        return personaObj;
    }

    Office.Controls.Persona.PersonaHelper.getPersonaObjectDef = function() {
        return { "id": "", "imgSrc": "", "primaryText": "", "secondaryText": "", "tertiaryText": "", "actions": { "email": "", "workPhone": "", "mobile" : "", "skype" : "" }};
    }

    Office.Controls.Persona.PersonaHelper.ensureJsonObjectLegal = function(legalObj, originObj) {
        if (typeof originObj !== 'object') {
            Office.Controls.Utils.errorConsole('illegal json object');
            return;
        }

        var key;
        for (key in legalObj) {
            if (typeof legalObj[key] === 'object') {
                Office.Controls.Persona.PersonaHelper.ensureJsonObjectLegal(legalObj[key], originObj[key]);
            } else {
                legalObj[key] = Office.Controls.Persona.StringUtils.setNullOrUndefinedAsEmpty(originObj[key]);
            }
        }
        return legalObj;
    }

    /**
     * Convert AAD User Data To Persona Object
     * @aadUserObject {JSON Object}
     * @return {JSON personaObject}
     */
    Office.Controls.Persona.PersonaHelper.convertAadUserToPersonaObject = function(aadUserObject, res) {
        if (typeof aadUserObject !== 'object' || (Office.Controls.Utils.isNullOrUndefined(aadUserObject))) {
            Office.Controls.Utils.errorConsole('AAD user data is null.');
            return;
        }

        var displayName = Office.Controls.Persona.StringUtils.getDisplayText(aadUserObject.displayName, Office.Controls.Persona.PersonaType.TypeEnum.NameImage, 0);
            
        var personaObj = {};
        personaObj.id = aadUserObject.id;
        personaObj.imgSrc = (!aadUserObject.imgSrc) ? Office.Controls.Persona.PersonaHelper._defaultImage: aadUserObject.imgSrc;
        personaObj.primaryText = (displayName === "") ? Office.Controls.Persona.PersonaHelper.getResourceString("NoName", res) : displayName;
        personaObj.secondaryText = "";

        if (!Office.Controls.Utils.isNullOrEmptyString(aadUserObject.jobTitle)) {
            personaObj.secondaryText = aadUserObject.jobTitle;
            if (!Office.Controls.Utils.isNullOrEmptyString(aadUserObject.department))
            {
               personaObj.secondaryText += Office.Controls.Persona.Strings.Comma + aadUserObject.department;
            }
        } else if (!Office.Controls.Utils.isNullOrEmptyString(aadUserObject.department)) {
            personaObj.secondaryText = aadUserObject.department;
        }

        personaObj.secondaryTextShort = Office.Controls.Persona.StringUtils.getDisplayText(personaObj.SecondaryText, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard, 2);
        personaObj.tertiaryText = Office.Controls.Persona.StringUtils.getDisplayText(aadUserObject.office, Office.Controls.Persona.PersonaType.TypeEnum.PersonaCard, 2);

        personaObj.actions = {};
        personaObj.actions.email = aadUserObject.mail;
        personaObj.actions.workPhone = aadUserObject.workPhone;
        personaObj.actions.mobile = aadUserObject.mobile;
        personaObj.actions.skype = aadUserObject.sipAddress;
        
        return personaObj;
    }

    Office.Controls.Persona.PersonaHelper.convertAadUsersToPersonaObjects = function(aadUsers, res) {
        if (typeof aadUsers !== 'object' || (Office.Controls.Utils.isNullOrUndefined(aadUsers))) {
            Office.Controls.Utils.errorConsole('AAD user collection is null.');
            return;
        }

        var personaObjects = [];
        aadUsers.forEach(function (aadUser) {
            personaObjects.push(Office.Controls.Persona.PersonaHelper.convertAadUserToPersonaObject(aadUser, res));
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
    
    Office.Controls.Persona.PersonaHelper.getResourceString = function (resName, res) {
        if (!Office.Controls.Utils.isNullOrUndefined(res)) {
            Office.Controls.Persona.PersonaHelper._resourceStrings = res;
        }
        
        // Check if the resource strings exsit
        if (Office.Controls.Persona.PersonaHelper._resourceStrings.hasOwnProperty(resName)) {
            return Office.Controls.Persona.PersonaHelper._resourceStrings[resName];    
        }
        return Office.Controls.Utils.getStringFromResource('Persona', resName);
    };
    
    Office.Controls.Persona.PersonaHelper.hideOpenedPersonaCard = function () {
        var showNodeQueue = Office.Controls.Persona.PersonaHelper._showNodeQueue;
        if (showNodeQueue.length !== 0) {
            var nodeItem = showNodeQueue.pop();
            nodeItem.showNode(nodeItem.get_rootNode(), false);
        }
    };

    // The Persona Type
    Office.Controls.Persona.PersonaType = function() {};
    Office.Controls.Persona.PersonaType.TypeEnum = {
        NameImage: "nameimage",
        PersonaCard: "personacard",
        ImageOnly: "imageonly"
    };

    Office.Controls.Persona.StringUtils = function () { };
    Office.Controls.Persona.StringUtils.getDisplayText = function (displayText, personaType, position) {
        if (!displayText) {
            return '';
        }
        
        // configurations of inline persona & cersonaCard
        var displayConfig = ((personaType === Office.Controls.Persona.PersonaType.TypeEnum.NameImage) ? [ 26, 26, 36, 42 ] : [ 18, 30, 32, 32 ]);
        
        var len = displayConfig[position];
        if (displayText.length > len) {
            return displayText.substr(0, len) + '...';
        } else {
            return displayText;
        }
    };

    Office.Controls.Persona.StringUtils.setNullOrUndefinedAsEmpty = function (str, value) {
        var val = ((value === undefined) ? "" : value);
        return Office.Controls.Utils.isNullOrEmptyString(str) ? val : str;
    };

    Office.Controls.Persona.Strings = function() {
    }

    Office.Controls.Persona.Templates = function() {
    }

    Office.Controls.PersonaResourcesDefaults = function () {
    };

    Office.Controls.PersonaConstants = function () {
    };

    
    if (Office.Controls.Persona.registerClass) { Office.Controls.Persona.registerClass('Office.Controls.Persona'); }
    if (Office.Controls.Persona.Strings.registerClass) { Office.Controls.Persona.Strings.registerClass('Office.Controls.Persona.Strings'); }
    if (Office.Controls.PersonaConstants.registerClass) { Office.Controls.PersonaConstants.registerClass('Office.Controls.PersonaConstants'); }
    if (Office.Controls.PersonaResourcesDefaults.registerClass) { Office.Controls.PersonaResourcesDefaults.registerClass('Office.Controls.PersonaResourcesDefaults'); }
    if (Office.Controls.Persona.PersonaHelper.registerClass) { Office.Controls.Persona.PersonaHelper.registerClass('Office.Controls.Persona.PersonaHelper'); }
    if (Office.Controls.Persona.StringUtils.registerClass) { Office.Controls.Persona.StringUtils.registerClass('Office.Controls.Persona.StringUtils'); }
    Office.Controls.PersonaResourcesDefaults.Email = 'Work: ';
    Office.Controls.PersonaResourcesDefaults.WorkPhone = 'Work: ';
    Office.Controls.PersonaResourcesDefaults.Mobile = 'Mobile: ';
    Office.Controls.PersonaResourcesDefaults.Skype = 'Skype: ';
    Office.Controls.PersonaResourcesDefaults.NoName = 'No Name';
    Office.Controls.Persona.Strings.Protocol_Mail = 'mailto:';
    Office.Controls.Persona.Strings.Protocol_Phone = 'tel:';
    Office.Controls.Persona.Strings.Protocol_Skype = 'sip:';
    Office.Controls.Persona.Strings.Comma = ', ';
    Office.Controls.Persona.Strings.SuspensionPoints = '...';
    Office.Controls.Persona.PersonaHelper._resourceStrings = {};
    Office.Controls.Persona.PersonaHelper._localCache = {};
    Office.Controls.Persona.PersonaHelper._showNodeQueue = [];
    Office.Controls.Persona.PersonaHelper._defaultImage = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAABGdBTUEAALGPC/xhBQAAAAFzUkdCAK7OHOkAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAORQTFRFsbGxv7+/wcHBu7u72dnZ7u7u/////Pz85ubmy8vLtbW1wsLC5OTk/f3939/ft7e34eHh+vr61NTUs7OzwMDA8/Pz7OzsxMTE9PT0vLy8+/v7+fn5urq6vb298fHxtLS0srKyzc3N9/f32NjYtra209PT7e3t7+/v/v7+yMjI3t7e4+Pj6+vr6urq4uLi19fXx8fH0tLS9fX1ubm59vb2z8/PuLi4+Pj429vbzs7O0NDQ4ODg1dXV2tra1tbWzMzMxcXF6enp0dHR8vLyysrKxsbG3d3d6Ojovr6+5eXl5+fn8PDwYCkYCwAAAAFiS0dEBmFmuH0AAAAJcEhZcwAAdTAAAHUwAd0zcs0AAAKgSURBVGje7ZjdVtpQEIUByxCIIRKtYECgSSVULbWopRDBarVq+/7vU1raBcTMmbEzueha7Ous/XEOZ35zuY022mij/0z5QqGQmfnWqyL8VskqV/Tt7W1YkVN1de0rO5BQzdtV9N97Dc+1r3eIegnS1DhQ8rd9SFezpeJ/2ARM7Y6Cf7cIuCwFgAcmvRH7t3wjIBCHXBXMCoX+rkMA/LcywBFQOpIBeiQgkt1QnwSAKNre0f6ylxoyAMcSwAkDsC8BnDIAgQQQMQCOBNBmAECSUotZA95zAJIrGjD8mxLABwagKAGcMQADCYCTiz5KAIxsCnkRwCP9hyJ/uqKJayb1UJ1zIeDc3FXAjtCfCoWa9ADzzv3CBLgU+8/HphLuf6rgn8uVUf/GJxUA2hwFOu37rzOk3tJQawCZy075p7eFTeO6OmHiECN5455Q63OwtG+PdWbM+qQaXx0uL2oa96IoGsxWE+h1bB1P7H9x3wpHi59bNQwZ7p+q3fZeOtR+Wan3DXQ/YS+vzYlf8mjd9Z7RuUr9ajdcS+ZOzE5Ml7Xkm7xJidn8s7bMP2PZd+OUqGom01onTKlF/VsOAJn7vpZXnmZrFqR/dUf7j9Hc5lu39UIlfz+Jh/g3ZA/g4psDlhrU5EwNxqSILuCA0WqZRUzOU6k/gPEldUdyQM8E4HS7lPqmtMSZW0mZSsWNBsA0OQdye+P6oqLhDz4O2FMBAJ6273QAdRRAzxssTVDANx3AFAVYOgB88OQsPxh6QAHCWvBX6GyuEwbzooMBHpUAaKTdKwEA61zHcuuFsMlnpgV4zDbO8FzxpAXA2i/Wio4jbHy+kFsv5CEAcrXCFdJ8uVr+WFXOqwFOMg5k+J5xIGMbebVAhh8ZBzK2KVQLZKituP4EqRB824c6sq4AAAAldEVYdGRhdGU6Y3JlYXRlADIwMTQtMTAtMjlUMjA6MjQ6MTktMDU6MDBCpOLkAAAAJXRFWHRkYXRlOm1vZGlmeQAyMDE0LTEwLTI5VDIwOjI0OjE5LTA1OjAwM/laWAAAAABJRU5ErkJggg==";
    Office.Controls.PersonaConstants.SectionTag_Main = "persona-section-tag-main";
    Office.Controls.PersonaConstants.SectionTag_Action = "ms-PersonaCard-action";
    Office.Controls.PersonaConstants.SectionTag_ActionDetail = "ms-PersonaCard-actionDetails";
    Office.Controls.Persona.Templates.DefaultDefinition = {
        "nameimage": 
        {
            value: "<div class=\"ms-Persona\"><div class=\"image\"><img class=\"imageOfNameImage\" style=\"background-image:url(${imgSrc})\"></img></div><div class=\"ms-Persona-details ms-Persona-details-nameImage\"><div class=\"ms-Persona-primaryText ms-Persona-primaryText-nameImage\"><Label class=\"clickStyle\" title=\"${primaryText}\">${primaryTextShort}</Label></div><div class=\"ms-Persona-secondaryText ms-Persona-secondaryText-nameImage\"><Label class=\"clickStyle\" title=\"${secondaryText}\">${secondaryTextShort}</Label></div></div></div>"
        },
        "personacard": 
        {
            value: "<div class=\"ms-PersonaCard personaCard-customized detail displayMode\"><div class=\"ms-PersonaCard-persona persona-section-tag-main\"><div class=\"ms-Persona ms-Persona--xl\"><div class=\"ms-Persona-imageArea\"><img class=\"ms-Persona-image image\" style=\"background-image:url(${imgSrc})\"></img></div><div class=\"ms-Persona-details\"><div class=\"ms-Persona-primaryText\"><Label class=\"defaultStyle\" title=\"${primaryText}\">${primaryTextShort}</Label></div><div class=\"ms-Persona-secondaryText\"><Label class=\"defaultStyle\" title=\"${secondaryText}\">${secondaryTextShort}</Label></div><div class=\"ms-Persona-tertiaryText\"><Label class=\"defaultStyle\" title=\"${tertiaryText}\">${tertiaryTextShort}</Label></div></div></div></div><ul class=\"ms-PersonaCard-actions\"><li class=\"ms-PersonaCard-action\" child=\"action-detail-mail\"><i class=\"ms-Icon ms-Icon--mail icon\"><span></span></i></li><li class=\"ms-PersonaCard-action\" child=\"action-detail-phone\"><i class=\"ms-Icon ms-Icon--phone icon\"><span></span></i></li><li class=\"ms-PersonaCard-action\" child=\"action-detail-chat\"><i class=\"ms-Icon ms-Icon--chat icon\"><span></span></i></li></ul><div class=\"ms-PersonaCard-actionDetails action-detail-mail\"><div class=\"ms-PersonaCard-detailLine\"><span class=\"ms-PersonaCard-detailLabel\">${strings.label.email}</span><a href=\"${strings.protocol.email}${actions.email}\">${actions.emailShort}</a></div></div><div class=\"ms-PersonaCard-actionDetails action-detail-phone\"><div class=\"ms-PersonaCard-detailLine\"><span class=\"ms-PersonaCard-detailLabel\">${strings.label.workPhone}</span><a href=\"${strings.protocol.phone}${actions.workPhone}\">${actions.workPhoneShort}</a><br/><span class=\"ms-PersonaCard-detailLabel\">${strings.label.mobile}</span><a href=\"${strings.protocol.phone}${actions.mobile}\">${actions.mobileShort}</a></div></div><div class=\"ms-PersonaCard-actionDetails action-detail-chat\"><div class=\"ms-PersonaCard-detailLine\"><span class=\"ms-PersonaCard-detailLabel\">${strings.label.skype}</span><a href=\"${strings.protocol.skype}${actions.skype}\">${actions.skypeShort}</a></div></div></div>"
        },
        "imageonly":
        {
            value: "<div class=\"ms-Persona ms-Persona--xs\"><div class=\"ms-Persona-imageArea\"><img class=\"ms-Persona-image\" src=\"${imgSrc}\"></img></div></div>"
        },
    };
})();