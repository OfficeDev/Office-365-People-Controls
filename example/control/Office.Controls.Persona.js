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
    *  {
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

        loadTemplateAsync: function (templatePath, callback)
        {
            if (Office.Controls.Utils.isNullOrUndefined(templatePath)) {
                Office.Controls.Utils.errorConsole('Wrong template path');
                return;
            }

            var self = this;
            var xmlhttp = new XMLHttpRequest();
            xmlhttp.open("GET", templatePath, true);
            xmlhttp.onreadystatechange = function() {
                // if (this.readyState !== 4) return;
                // if (this.status !== 200) return; // or whatever error handling you want
                if (this.readyState === 4) {
                    if (this.status === 200) {
                        var parser, xmlDoc
                        if (window.DOMParser) {
                           parser = new DOMParser();
                           xmlDoc = parser.parseFromString(this.responseText,"text/xml");
                        } else { // code for < IE9
                           xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
                           xmlDoc.async = false;
                           xmlDoc.loadXML(this.responseText); 
                        }  
                        self.parseTemplate(xmlDoc);
                        callback(self.rootNode, null); 
                    }
                }
            };
            xmlhttp.send();
        },

        parseTemplate: function (xmlDoc)
        {
            try {
                if (typeof xmlDoc !== 'object' || (Office.Controls.Utils.isNullOrUndefined(xmlDoc))) {
                    Office.Controls.Utils.errorConsole('Invalid template document');
                    return;
                }
                
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

        bindData: function (htmlStr)
        {
            var regExp = /\$\{([^\}\{]+)\}/g;
            var resultStr = htmlStr;
            // Get the data
            var displayInfo = this.dataProvider;

            // Get the property names
            var properties = resultStr.match(regExp);

            for (var i = 0; i < properties.length; i++) { 
                var propertyValue = displayInfo[properties[i].substring(2, properties[i].length - 1)]
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
        },

        setActiveStyle: function (event)
        {
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

    Office.Controls.Persona.Strings = function() {
    }

    Office.Controls.PersonaResources = function () {
    };

    Office.Controls.PersonaConstants = function () {
    };

    if (Office.Controls.Persona.registerClass) { Office.Controls.Persona.registerClass('Office.Controls.Persona'); }
    // if (Office.Controls.Persona.PersonaViewModel.registerClass) {Office.Controls.Persona.PersonaViewModel.registerClass('Office.Controls.Persona.PersonaViewModel');}
    if (Office.Controls.Persona.Strings.registerClass) { Office.Controls.Persona.Strings.registerClass('Office.Controls.Persona.Strings'); }
    if (Office.Controls.PersonaConstants.registerClass) { Office.Controls.PersonaConstants.registerClass('Office.Controls.PersonaConstants'); }
    if (Office.Controls.PersonaResources.registerClass) { Office.Controls.PersonaResources.registerClass('Office.Controls.PersonaResources'); }
    Office.Controls.PersonaResources.PersonaName = 'Persona';
    Office.Controls.Persona.Strings.emailString = 'Email';
    Office.Controls.Persona.Strings.lyncMessageString = 'IM';
    Office.Controls.Persona.Strings.phoneString = 'Phone';
    Office.Controls.Persona.Strings.mobileString = 'Mobile';
    Office.Controls.Persona.Strings.workPhoneString = 'Work';
    Office.Controls.Persona.Strings.colonString = ':';
    Office.Controls.Persona.Strings.suspensionPoints = '...';
    Office.Controls.Persona.res = {};
    Office.Controls.PersonaConstants.SectionTag_Main = "persona-section-tag-main";
    Office.Controls.PersonaConstants.SectionTag_Action = "ms-PersonaCard-action";
    Office.Controls.PersonaConstants.SectionTag_ActionDetail = "ms-PersonaCard-actionDetails";
})();