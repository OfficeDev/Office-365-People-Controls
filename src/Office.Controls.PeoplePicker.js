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

    Office.Controls.PrincipalInfo = function () { };

    Office.Controls.PeoplePickerRecord = function () { };

    Office.Controls.PeoplePickerRecord.prototype = {
        isResolved: false,
        text: null,
        displayName: null,
        description: null,
        id: null,
        imgSrc: null,
        principalInfo: null
    }

    Office.Controls.PeoplePicker = function (root, dataProvider, options) {
        try {
            if (typeof root !== 'object' || typeof dataProvider !== 'object' || (!Office.Controls.Utils.isNullOrUndefined(options) && typeof options !== 'object')) {
                Office.Controls.Utils.errorConsole('Invalid parameters type');
                return;
            }
            this.root = root;
            this.dataProvider = dataProvider;
            if (!Office.Controls.Utils.isNullOrUndefined(options)) {
                if (!Office.Controls.Utils.isNullOrUndefined(options.allowMultipleSelections)) {
                    this.allowMultipleSelections = (String(options.allowMultipleSelections) === "true");
                }
                if (!Office.Controls.Utils.isNullOrUndefined(options.startSearchCharLength) && options.startSearchCharLength >= 1) {
                    this.startSearchCharLength = options.startSearchCharLength;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(options.delaySearchInterval)) {
                    this.delaySearchInterval = options.delaySearchInterval;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(options.enableCache)) {
                    this.enableCache = (String(options.enableCache) === "true");
                }
                if (!Office.Controls.Utils.isNullOrUndefined(options.numberOfResults)) {
                    this.numberOfResults = options.numberOfResults;
                }
                if (!Office.Controls.Utils.isNullOrEmptyString(options.inputHint)) {
                    this.inputHint = options.inputHint;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(options.showValidationErrors)) {
                    this.showValidationErrors = (String(options.showValidationErrors) === "true");
                }
                if (!Office.Controls.Utils.isNullOrUndefined(options.showImage)) {
                    this.showImage = (String(options.showImage) === "true");
                }
                if (!Office.Controls.Utils.isNullOrUndefined(options.onAdd)) {
                    this.onAdd = options.onAdd;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(options.onRemove)) {
                    this.onRemove = options.onRemove;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(options.onChange)) {
                    this.onChange = options.onChange;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(options.onFocus)) {
                    this.onFocus = options.onFocus;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(options.onBlur)) {
                    this.onBlur = options.onBlur;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(options.hideResultRecord)) {
                    this.hideResultRecord = (String(options.hideResultRecord) === "true");
                }

                this.onError = options.onError;

                if (!Office.Controls.Utils.isNullOrUndefined(options.resourceStrings)) {
                    Office.Controls.PeoplePicker.resourceStrings = options.resourceStrings;
                }
            }

            this.currentTimerId = -1;
            this.selectedItems = new Array(0);
            this.internalSelectedItems = new Array(0);
            this.errors = new Array(0);
            if (this.enableCache === true) {
                Office.Controls.Runtime.initialize({ HostUrl: window.location.host });
                this.cache = Office.Controls.PeoplePicker.mruCache.getInstance();
            }

            this.renderControl();
            this.autofill = new Office.Controls.PeoplePicker.autofillContainer(this);
        } catch (ex) {
            throw ex;
        }
    };

    Office.Controls.PeoplePicker.copyToRecord = function (record, info) {
        record.displayName = info.displayName;
        record.description = info.description;
        record.id = info.id;
        record.imgSrc = info.imgSrc;
        record.principalInfo = info;
    };

    Office.Controls.PeoplePicker.parseUserPaste = function (content) {
        var openBracket = content.indexOf('<'), emailSep = content.indexOf('@', openBracket),
        closeBracket = content.indexOf('>', emailSep);
        if (openBracket !== -1 && emailSep !== -1 && closeBracket !== -1) {
            return content.substring(openBracket + 1, closeBracket);
        }
        return content;
    };
    Office.Controls.PeoplePicker.getSearchBoxClass = function () {
        return 'ms-PeoplePicker-searchBox';
    };
    Office.Controls.PeoplePicker.nopAddRemove = function () { };
    Office.Controls.PeoplePicker.nopOperation = function () { };

    Office.Controls.PeoplePicker.create = function (root, contextOrGetTokenAsync, options) {
        var dataProvider = new Office.Controls.PeopleAadDataProvider(contextOrGetTokenAsync);
        return new Office.Controls.PeoplePicker(root, dataProvider, options);
    };
    Office.Controls.PeoplePicker.prototype = {
        allowMultipleSelections: false,
        startSearchCharLength: 1,
        delaySearchInterval: 300,
        enableCache: true,
        inputHint: null,
        numberOfResults: 30,
        onAdd: Office.Controls.PeoplePicker.nopAddRemove,
        onRemove: Office.Controls.PeoplePicker.nopAddRemove,
        onChange: Office.Controls.PeoplePicker.nopOperation,
        onFocus: Office.Controls.PeoplePicker.nopOperation,
        onBlur: Office.Controls.PeoplePicker.nopOperation,
        onError: null,
        dataProvider: null,
        showValidationErrors: true,
        showImage: true,
        showInputHint: true,
        inputTabindex: 0,
        searchingTimes: 0,
        inputBeginAction: false,
        actualRoot: null,
        textInput: null,
        inputData: null,
        defaultText: null,
        resolvedListRoot: null,
        autofillElement: null,
        errorMessageElement: null,
        root: null,
        lastSearchQuery: '',
        currentToken: null,
        widthSet: false,
        currentPrincipalsChoices: null,
        hasErrors: false,
        errorDisplayed: null,
        hasMultipleMatchValidationError: false,
        hasNoMatchValidationError: false,
        autofill: null,
        hideResultRecord: false,

        reset: function () {
            var record;
            while (this.internalSelectedItems.length) {
                record = this.internalSelectedItems[0];
                record.removeAndNotTriggerUserListener();
            }
            this.setTextInputDisplayStyle();
            this.validateMultipleMatchError();
            this.validateNoMatchError();
            this.clearInputField();
            this.clearCacheData();
            if (Office.Controls.PeoplePicker.autofillContainer.currentOpened) {
                Office.Controls.PeoplePicker.autofillContainer.currentOpened.close();
            }
            Office.Controls.PeoplePicker.autofillContainer.currentOpened = null;
            Office.Controls.PeoplePicker.autofillContainer.boolBodyHandlerAdded = false;
            this.autofill = new Office.Controls.PeoplePicker.autofillContainer(this);
            this.toggleDefaultText();

        },

        remove: function (entryToRemove) {
            var record = this.internalSelectedItems, i, recordToRemove;
            for (i = 0; i < record.length; i++) {
                if (record[i].Record.principalInfo === entryToRemove) {
                    recordToRemove = record[i].Record;
                    record[i].removeAndNotTriggerUserListener();
                    this.onRemove(this, recordToRemove.principalInfo);
                    this.validateMultipleMatchError();
                    this.validateNoMatchError();
                    this.setTextInputDisplayStyle();
                    this.textInput.focus();
                    break;
                }
            }
        },

        add: function (p1, resolved) {
            if (typeof p1 === 'string') {
                this.addThroughString(p1);
            } else {
                var record = new Office.Controls.PeoplePickerRecord();
                Office.Controls.PeoplePicker.copyToRecord(record, p1);
                record.text = p1.displayName;
                if (Office.Controls.Utils.isNullOrUndefined(resolved)) {
                    record.isResolved = false;
                    this.addThroughRecord(record, false);
                } else {
                    record.isResolved = resolved;
                    this.addThroughRecord(record, resolved);
                }
            }
        },

        getAddedPeople: function () {
            var record = this.internalSelectedItems, addedPeople = [], i;
            for (i = 0; i < record.length; i++) {
                addedPeople[i] = record[i].Record.principalInfo;
            }
            return addedPeople;
        },

        clearCacheData: function () {
            if (this.cache !== null) {
                this.cache.cacheDelete('Office.PeoplePicker.Cache');
                this.cache.dataObject = new Office.Controls.PeoplePicker.mruCache.mruData();
                this.cache.dataObject.cacheMapping[Office.Controls.Runtime.context.HostUrl] = new Array(0);
            }
        },

        getErrorDisplayed: function () {
            return this.errorDisplayed;
        },

        getUserInfoAsync: function (userInfoHandler, userEmail) {
            var record = new Office.Controls.PeoplePickerRecord();
            this.dataProvider.searchPeopleAsync(userEmail, function (error, principalsReceived) {
                if (principalsReceived !== null) {
                    Office.Controls.PeoplePicker.copyToRecord(record, principalsReceived[0]);
                    userInfoHandler(record);
                } else {
                    userInfoHandler(null);
                }
            });
        },

        get_textInput: function () {
            return this.textInput;
        },

        get_actualRoot: function () {
            return this.actualRoot;
        },

        addThroughString: function (input) {
            if (Office.Controls.Utils.isNullOrEmptyString(input)) {
                Office.Controls.Utils.errorConsole('Input can\'t be null or empty string. PeoplePicker Id : ' + this.root.id);
                return;
            }
            this.addUnresolvedPrincipal(input, false);
        },

        addThroughRecord: function (info, resolved) {
            if (!resolved) {
                this.addUncertainPrincipal(info);
            } else {
                this.addResolvedRecord(info);
            }
        },

        renderControl: function (inputName) {
            this.root.innerHTML = Office.Controls.peoplePickerTemplates.generateControlTemplate(inputName, this.allowMultipleSelections, this.inputHint);
            if (this.root.className.length > 0) {
                this.root.className += ' ';
            }
            this.root.className += 'office office-peoplepicker';
            this.actualRoot = this.root.querySelector('div.ms-PeoplePicker');
            var self = this;
            Office.Controls.Utils.addEventListener(this.actualRoot, 'click', function (e) {
                return self.onPickerClick(e);
            });
            this.inputData = this.actualRoot.querySelector('input[type=\"hidden\"]');
            this.textInput = this.actualRoot.querySelector('input[type=\"text\"]');
            this.defaultText = this.actualRoot.querySelector('span.office-peoplepicker-default');
            this.resolvedListRoot = this.actualRoot.querySelector('div.office-peoplepicker-recordList');
            this.autofillElement = this.actualRoot.querySelector('.ms-PeoplePicker-results');
            Office.Controls.Utils.addEventListener(this.textInput, 'focus', function (e) {
                return self.onInputFocus(e);
            });
            Office.Controls.Utils.addEventListener(this.textInput, 'blur', function (e) {
                return self.onInputBlur(e);
            });
            Office.Controls.Utils.addEventListener(this.textInput, 'keydown', function (e) {
                return self.onInputKeyDown(e);
            });
            Office.Controls.Utils.addEventListener(this.textInput, 'input', function (e) {
                return self.onInput(e);
            });
            Office.Controls.Utils.addEventListener(window.self, 'resize', function (e) {
                return self.onResize(e);
            });
            this.toggleDefaultText();
            if (!Office.Controls.Utils.isNullOrUndefined(this.inputTabindex)) {
                this.textInput.setAttribute('tabindex', this.inputTabindex);
            }
        },

        toggleDefaultText: function () {
            if (this.actualRoot.className.indexOf('office-peoplepicker-autofill-focus') === -1 && this.showInputHint && !this.selectedItems.length && !this.textInput.value.length) {
                this.defaultText.className = 'office-peoplepicker-default';
            } else {
                this.defaultText.className = 'office-hide';
            }
        },

        onResize: function () {
            this.toggleDefaultText();
            return true;
        },

        onInputKeyDown: function (e) {
            var keyEvent = Office.Controls.Utils.getEvent(e), self = this;
            if (keyEvent.keyCode === 27) { // 'escape'
                this.autofill.close();
            } else if ((keyEvent.keyCode === 9 || keyEvent.keyCode === 13) && this.autofill.IsDisplayed) { // 'tab' || 'enter'
                var focusElement = this.autofillElement.querySelector("li.ms-PeoplePicker-resultAddedForSelect");
                if (focusElement !== null) {
                    var personId = this.autofill.getPersonIdFromListElement(focusElement);
                    this.addResolvedPrincipal(this.autofill.entries[personId]);
                    this.autofill.flushContent();
                    Office.Controls.Utils.cancelEvent(e);
                    return false;
                }
                this.autofill.close();
            } else if ((keyEvent.keyCode === 40 || keyEvent.keyCode === 38) && this.autofill.IsDisplayed) { // 'down arrow' || 'up arrow'
                this.autofill.onKeyDownFromInput(keyEvent);
                keyEvent.preventDefault();
                keyEvent.stopPropagation();
                Office.Controls.Utils.cancelEvent(e);
                return false;
            } else if (keyEvent.keyCode === 37 && this.internalSelectedItems.length) { // 'left arrow'
                this.resolvedListRoot.lastChild.focus();
                Office.Controls.Utils.cancelEvent(e);
                return false;
            } else if (keyEvent.keyCode === 8) { // 'backspace'
                var shouldRemove = false;
                if (!Office.Controls.Utils.isNullOrUndefined(document.selection)) {
                    var range = document.selection.createRange(), selectedText = range.text, caretPos = range.text.length;
                    range.moveStart('character', -this.textInput.value.length);
                    if (!selectedText.length && !caretPos) {
                        shouldRemove = true;
                    }
                } else {
                    var selectionStart = this.textInput.selectionStart, selectionEnd = this.textInput.selectionEnd;
                    if (!selectionStart && selectionStart === selectionEnd) {
                        shouldRemove = true;
                    }
                }
                if (shouldRemove && this.internalSelectedItems.length) {
                    this.internalSelectedItems[this.internalSelectedItems.length - 1].remove();
                    Office.Controls.Utils.cancelEvent(e);
                }
            } else if ((keyEvent.keyCode === 75 && keyEvent.ctrlKey) || (keyEvent.keyCode === 186) || (keyEvent.keyCode === 59) || (keyEvent.keyCode === 13)) {
                // 'Ctrl + k' || 'semi-colon' (59 on Firefox and 186 on other browsers) || 'enter'
                keyEvent.preventDefault();
                keyEvent.stopPropagation();
                this.cancelLastRequest();
                if (!this.hideResultRecord) {
                    this.attemptResolveInput();
                }
                Office.Controls.Utils.cancelEvent(e);
                return false;
            } else if ((keyEvent.keyCode === 86 && keyEvent.ctrlKey) || (keyEvent.keyCode === 186)) {
                // 'Ctrl + v' || 'semi-colon'
                this.cancelLastRequest();
                window.setTimeout(function () {
                    self.textInput.value = Office.Controls.PeoplePicker.parseUserPaste(self.textInput.value);
                    if (!self.hideResultRecord) {
                        self.attemptResolveInput();
                    }
                }, 0);
                return true;
            } else if (keyEvent.keyCode === 13 && keyEvent.shiftKey) { // 'shift + enter'
                this.autofill.open(function (selectedPrincipal) {
                    self.addResolvedPrincipal(selectedPrincipal);
                });
            } else {
                this.resizeInputField();
            }
            return true;
        },

        cancelLastRequest: function () {
            window.clearTimeout(this.currentTimerId);
            if (!Office.Controls.Utils.isNullOrUndefined(this.currentToken)) {
                this.hideLoadingIcon();
                this.currentToken.cancel();
                this.currentToken = null;
            }
        },

        onInput: function (e) {
            this.startQueryAfterDelay();
            this.resizeInputField();
            this.autofill.close();
            return true;
        },

        displayCachedEntries: function () {
            var cachedEntries = this.cache.get(this.textInput.value, 5), self = this;
            this.autofill.setCachedEntries(cachedEntries);
            this.autofill.setServerEntries(new Array(0));
            if (!cachedEntries.length) {
                return;
            }
            this.autofill.open(function (selectedPrincipal) {
                self.addResolvedPrincipal(selectedPrincipal);
            });
        },

        resizeInputField: function () {
            var size = Math.max(this.textInput.value.length + 1, 1);
            this.textInput.size = size;
        },

        clearInputField: function () {
            this.textInput.value = '';
            this.resizeInputField();
        },

        startQueryAfterDelay: function () {
            this.cancelLastRequest();
            var currentValue = this.textInput.value, self = this;
            this.currentTimerId = window.setTimeout(function () {
                if (currentValue !== self.lastSearchQuery || self.startSearchCharLength === 0) {
                    self.lastSearchQuery = currentValue;
                    if (currentValue.length >= self.startSearchCharLength) {
                        self.searchingTimes++;
                        self.removeValidationError('ServerProblem');
                        if (self.enableCache) {
                            self.displayCachedEntries();
                            self.displayLoadingIcon(currentValue, true);
                        } else {
                            self.displayLoadingIcon(currentValue, false);
                        }
                        var token = new Office.Controls.PeoplePicker.cancelToken();
                        self.currentToken = token;
                        self.dataProvider.searchPeopleAsync(self.textInput.value, function (error, principalsReceived) {
                            if (!token.IsCanceled) {
                                if (principalsReceived !== null) {
                                    self.onDataReceived(principalsReceived);
                                } else {
                                    self.onDataFetchError(error);
                                }
                            }
                        });
                    } else {
                        self.autofill.close();
                        if (self.enableCache) {
                            self.displayCachedEntries();
                        }
                    }
                }
            }, self.delaySearchInterval);
        },

        onDataFetchError: function (message) {
            this.hideLoadingIcon();
            this.addValidationError(Office.Controls.PeoplePicker.ValidationError.createServerProblemError());
        },

        onDataReceived: function (principalsReceived) {
            if (! Office.Controls.Utils.isNullOrUndefined(this.textInput.value ) && this.textInput.value !== " " && this.textInput.value.length>= this.startSearchCharLength) {
                this.currentPrincipalsChoices = {};
                var self = this, i;
                for (i = 0; i < principalsReceived.length; i++) {
                    var principal = principalsReceived[i];
                    this.currentPrincipalsChoices[principal.id] = principal;
                }
                this.autofill.setServerEntries(principalsReceived);
                this.hideLoadingIcon();
                this.autofill.open(function (selectedPrincipal) {
                    self.addResolvedPrincipal(selectedPrincipal);
                });
            }
        },

        onPickerClick: function (e) {
            this.textInput.focus();
            e = Office.Controls.Utils.getEvent(e);
            var element = Office.Controls.Utils.getTarget(e);
            if (element.nodeName.toLowerCase() !== 'input') {
                this.focusToEnd();
            }
            return true;
        },

        focusToEnd: function () {
            var endPos = this.textInput.value.length;
            if (!Office.Controls.Utils.isNullOrUndefined(this.textInput.createTextRange)) {
                var range = this.textInput.createTextRange();
                range.collapse(true);
                range.moveStart('character', endPos);
                range.moveEnd('character', endPos);
                range.select();
            } else {
                this.textInput.focus();
                this.textInput.setSelectionRange(endPos, endPos);
            }
        },

        onInputFocus: function (e) {
            var self = this;
            if (Office.Controls.Utils.isNullOrEmptyString(this.actualRoot.className)) {
                this.actualRoot.className = 'office-peoplepicker-autofill-focus';
            } else {
                this.actualRoot.className += ' office-peoplepicker-autofill-focus';
            }
            if (!this.widthSet) {
                this.setInputMaxWidth();
            }
            this.toggleDefaultText();
            this.onFocus(this);

            if (this.startSearchCharLength === 0 && (this.allowMultipleSelections === true || this.internalSelectedItems.length === 0)) {
                self.startQueryAfterDelay();
            }
            return true;
        },

        setInputMaxWidth: function () {
            var maxwidth = this.actualRoot.clientWidth - 25;
            if (maxwidth <= 0) {
                maxwidth = 20;
            }
            this.textInput.style.maxWidth = maxwidth.toString() + 'px';
            this.widthSet = true;
        },

        onInputBlur: function (e) {
            Office.Controls.Utils.removeClass(this.actualRoot, 'office-peoplepicker-autofill-focus');
            if (this.textInput.value.length > 0 || this.selectedItems.length > 0) {
                this.onBlur(this);
                return true;
            }
            this.toggleDefaultText();
            this.onBlur(this);
            return true;
        },

        onDataSelected: function (selectedPrincipal) {
            this.lastSearchQuery = '';
            this.clearInputField();
            this.refreshInputField();
        },

        onDataRemoved: function (selectedPrincipal) {
            this.refreshInputField();
            this.validateMultipleMatchError();
            this.validateNoMatchError();
            this.onRemove(this, selectedPrincipal.principalInfo);
            this.onChange(this);
        },

        addToCache: function (entry) {
            if (!this.cache.isCacheAvailable) {
                return;
            }
            this.cache.set(entry);
        },

        refreshInputField: function () {
            this.inputData.value = Office.Controls.Utils.serializeJSON(this.selectedItems);
            this.setTextInputDisplayStyle();
        },

        setTextInputDisplayStyle: function () {
            if ((!this.allowMultipleSelections) && (this.internalSelectedItems.length === 1)) {
                this.textInput.className = 'ms-PeoplePicker-searchFieldAddedForSingleSelectionHidden';
                this.textInput.setAttribute('readonly', 'readonly');
            } else {
                this.textInput.removeAttribute('readonly');
                this.textInput.className = 'ms-PeoplePicker-searchField ms-PeoplePicker-searchFieldAdded';
            }
        },

        displayLoadingIcon: function (searchingName, isAppended) {
            this.autofill.openSearchingLoadingStatus(searchingName, isAppended);
        },

        hideLoadingIcon: function () {
            this.autofill.closeSearchingLoadingStatus();
        },

        attemptResolveInput: function () {
            this.autofill.close();
            if (this.textInput.value.length > 0) {
                this.lastSearchQuery = '';
                this.addUnresolvedPrincipal(this.textInput.value, true);
            }
        },

        onDataReceivedForResolve: function (principalsReceived, internalRecordToResolve) {
            this.hideLoadingIcon();
            if (principalsReceived.length === 1) {
                internalRecordToResolve.resolveTo(principalsReceived[0]);
            } else {
                internalRecordToResolve.setResolveOptions(principalsReceived);
            }
            this.refreshInputField();
            return internalRecordToResolve;
        },

        onDataReceivedForStalenessCheck: function (principalsReceived, internalRecordToCheck) {
            if (principalsReceived.length === 1) {
                internalRecordToCheck.resolveTo(principalsReceived[0]);
            } else {
                internalRecordToCheck.unresolve();
                internalRecordToCheck.setResolveOptions(principalsReceived);
            }
            this.refreshInputField();
        },

        addResolvedPrincipal: function (principal) {
            var record = new Office.Controls.PeoplePickerRecord(),
            internalRecord = new Office.Controls.PeoplePicker.internalPeoplePickerRecord(this, record);
            Office.Controls.PeoplePicker.copyToRecord(record, principal);
            record.text = principal.displayName;
            record.isResolved = true;
            this.selectedItems.push(record);
            internalRecord.add();
            internalRecord.updateHoverText();
            this.internalSelectedItems.push(internalRecord);
            this.onDataSelected(record);
            if (this.enableCache) {
                this.addToCache(principal);
            }
            this.currentPrincipalsChoices = null;
            this.autofill.close();
            this.textInput.focus();
            this.onAdd(this, record.principalInfo);
            this.onChange(this);
        },

        addResolvedRecord: function (record) {
            this.selectedItems.push(record);
            var internalRecord = new Office.Controls.PeoplePicker.internalPeoplePickerRecord(this, record);
            internalRecord.add();
            internalRecord.updateHoverText();
            this.internalSelectedItems.push(internalRecord);
            this.onDataSelected(record);
            this.onAdd(this, record.principalInfo);
            this.currentPrincipalsChoices = null;
        },

        addUncertainPrincipal: function (record) {
            this.selectedItems.push(record);
            var internalRecord = new Office.Controls.PeoplePicker.internalPeoplePickerRecord(this, record),
            self = this;
            internalRecord.add();
            internalRecord.updateHoverText();
            this.internalSelectedItems.push(internalRecord);
            this.setTextInputDisplayStyle();
            this.displayLoadingIcon(record.text, false);
            this.dataProvider.searchPeopleAsync(record.displayName, function (error, ps) {
                if (ps !== null) {
                    internalRecord = self.onDataReceivedForResolve(ps, internalRecord);
                    self.onAdd(this, internalRecord.Record.principalInfo);
                    self.onChange(self);
                } else {
                    self.onDataFetchError(error);
                }
            });
        },

        addUnresolvedPrincipal: function (input, triggerUserListener) {
            var record = new Office.Controls.PeoplePickerRecord(), self = this,
            principalInfo = new Office.Controls.PrincipalInfo();
            principalInfo.displayName = input;
            record.text = input;
            record.principalInfo = principalInfo;
            record.isResolved = false;
            var internalRecord = new Office.Controls.PeoplePicker.internalPeoplePickerRecord(this, record);
            internalRecord.add();
            internalRecord.updateHoverText();
            this.selectedItems.push(record);
            this.internalSelectedItems.push(internalRecord);
            this.clearInputField();
            this.setTextInputDisplayStyle();
            this.displayLoadingIcon(input, false);
            this.dataProvider.searchPeopleAsync(input, function (error, ps) {
                if (ps !== null) {
                    internalRecord = self.onDataReceivedForResolve(ps, internalRecord);
                    if (triggerUserListener) {
                        self.onAdd(self, internalRecord.Record.principalInfo);
                        self.onChange(self);
                    }
                } else {
                    self.onDataFetchError(error);
                }
            });
        },

        addValidationError: function (err) {
            this.hasErrors = true;
            this.errors.push(err);
            this.displayValidationErrors();
            if (!Office.Controls.Utils.isNullOrUndefined(this.onError)) {
                this.onError(this, err);
            }
        },

        removeValidationError: function (errorName) {
            var i;
            for (i = 0; i < this.errors.length; i++) {
                if (this.errors[i].errorName === errorName) {
                    this.errors.splice(i, 1);
                    break;
                }
            }
            if (!this.errors.length) {
                this.hasErrors = false;
            }
            if (!Office.Controls.Utils.isNullOrUndefined(this.onError) && this.errors.length) {
                this.onError(this, this.errors[0]);
            } else {
                this.displayValidationErrors();
            }
        },

        validateMultipleMatchError: function () {
            var oldStatus = this.hasMultipleMatchValidationError, newStatus = false, i;
            for (i = 0; i < this.internalSelectedItems.length; i++) {
                if (!Office.Controls.Utils.isNullOrUndefined(this.internalSelectedItems[i].optionsList) && this.internalSelectedItems[i].optionsList.length > 0) {
                    newStatus = true;
                    break;
                }
            }
            if (!oldStatus && newStatus) {
                this.addValidationError(Office.Controls.PeoplePicker.ValidationError.createMultipleMatchError());
            }
            if (oldStatus && !newStatus) {
                this.removeValidationError('MultipleMatch');
            }
            this.hasMultipleMatchValidationError = newStatus;
        },

        validateNoMatchError: function () {
            var oldStatus = this.hasNoMatchValidationError, newStatus = false, i;
            for (i = 0; i < this.internalSelectedItems.length; i++) {
                if (!Office.Controls.Utils.isNullOrUndefined(this.internalSelectedItems[i].optionsList) && !this.internalSelectedItems[i].optionsList.length) {
                    newStatus = true;
                    break;
                }
            }
            if (!oldStatus && newStatus) {
                this.addValidationError(Office.Controls.PeoplePicker.ValidationError.createNoMatchError());
            }
            if (oldStatus && !newStatus) {
                this.removeValidationError('NoMatch');
            }
            this.hasNoMatchValidationError = newStatus;
        },

        displayValidationErrors: function () {
            if (!this.showValidationErrors) {
                return;
            }
            if (!this.errors.length) {
                if (!Office.Controls.Utils.isNullOrUndefined(this.errorMessageElement)) {
                    this.errorMessageElement.parentNode.removeChild(this.errorMessageElement);
                    this.errorMessageElement = null;
                    this.errorDisplayed = null;
                }
            } else {
                if (this.errorDisplayed !== this.errors[0]) {
                    if (!Office.Controls.Utils.isNullOrUndefined(this.errorMessageElement)) {
                        this.errorMessageElement.parentNode.removeChild(this.errorMessageElement);
                    }
                    var holderDiv = document.createElement('div');
                    holderDiv.innerHTML = Office.Controls.peoplePickerTemplates.generateErrorTemplate(this.errors[0].localizedErrorMessage);
                    this.errorMessageElement = holderDiv.firstChild;
                    this.root.appendChild(this.errorMessageElement);
                    this.errorDisplayed = this.errors[0];
                }
            }
        },

        setDataProvider: function (newProvider) {
            this.dataProvider = newProvider;
        }
    };

    Office.Controls.PeoplePicker.internalPeoplePickerRecord = function (parent, record) {
        this.parent = parent;
        this.Record = record;
    };

    Office.Controls.PeoplePicker.internalPeoplePickerRecord.prototype = {
        Record: null,

        get_record: function () {
            return this.Record;
        },

        set_record: function (value) {
            this.Record = value;
            return value;
        },

        _principalOptions: null,
        _optionsList: null,
        Node: null,

        get_node: function () {
            return this.Node;
        },

        set_node: function (value) {
            this.Node = value;
            return value;
        },

        parent: null,

        onRecordRemovalClick: function (e) {
            var recordRemovalEvent = Office.Controls.Utils.getEvent(e),
            target = Office.Controls.Utils.getTarget(recordRemovalEvent);
            this.remove();
            Office.Controls.Utils.cancelEvent(e);
            this.parent.autofill.close();
            return false;
        },

        onRecordRemovalKeyDown: function (e) {
            var recordRemovalEvent = Office.Controls.Utils.getEvent(e);
            if (recordRemovalEvent.keyCode === 8 || recordRemovalEvent.keyCode === 13 || recordRemovalEvent.keyCode === 46) {
                this.remove();
                Office.Controls.Utils.cancelEvent(e);
                this.parent.autofill.close();
            }
            return false;
        },

        onRecordKeyDown: function (e) {
            var keyEvent = Office.Controls.Utils.getEvent(e);
            var target = Office.Controls.Utils.getTarget(keyEvent);
            if (keyEvent.keyCode === 8 || keyEvent.keyCode === 13 || keyEvent.keyCode === 46) {
                this.remove();
                Office.Controls.Utils.cancelEvent(e);
                this.parent.autofill.close();
            } else if (keyEvent.keyCode === 37) {
                if (this.Node.previousSibling !== null) {
                    this.Node.previousSibling.focus();
                }
                Office.Controls.Utils.cancelEvent(e);
            } else if (keyEvent.keyCode === 39) {
                if (this.Node.nextSibling !== null) {
                    this.Node.nextSibling.focus();
                } else {
                    this.parent.textInput.focus();
                }
                Office.Controls.Utils.cancelEvent(e);
            }
            return false;
        },

        add: function () {
            var holderDiv = document.createElement('div');
            holderDiv.innerHTML = Office.Controls.peoplePickerTemplates.generateRecordTemplate(this.Record, this.parent.allowMultipleSelections, this.parent.showImage);
            var recordElement = holderDiv.firstChild,
            removeButtonElement = recordElement.querySelector('div.ms-PeoplePicker-personaRemove'),
            self = this;
            Office.Controls.Utils.addEventListener(recordElement, 'keydown', function (e) {
                return self.onRecordKeyDown(e);
            });

            Office.Controls.Utils.addEventListener(removeButtonElement, 'click', function (e) {
                return self.onRecordRemovalClick(e);
            });
            Office.Controls.Utils.addEventListener(removeButtonElement, 'keydown', function (e) {
                return self.onRecordRemovalKeyDown(e);
            });
            
            this.parent.resolvedListRoot.appendChild(recordElement);

            if (this.parent.hideResultRecord) {
                recordElement.className = 'office-hide';
            }
            else
            {
                this.parent.defaultText.className = 'office-hide';
            }
            
            this.Node = recordElement;
        },

        remove: function () {
            this.removeAndNotTriggerUserListener();
            this.parent.onDataRemoved(this.Record);
            this.parent.textInput.focus();
        },

        removeAndNotTriggerUserListener: function () {
            this.parent.resolvedListRoot.removeChild(this.Node);
            var i;
            for (i = 0; i < this.parent.internalSelectedItems.length; i++) {
                if (this.parent.internalSelectedItems[i] === this) {
                    this.parent.internalSelectedItems.splice(i, 1);
                }
            }
            for (i = 0; i < this.parent.selectedItems.length; i++) {
                if (this.parent.selectedItems[i] === this.Record) {
                    this.parent.selectedItems.splice(i, 1);
                }
            }
        },

        setResolveOptions: function (options) {
            this.optionsList = options;
            this.principalOptions = {};
            var i;
            for (i = 0; i < options.length; i++) {
                this.principalOptions[options[i].id] = options[i];
            }
            var self = this;
            Office.Controls.Utils.addEventListener(this.Node, 'click', function (e) {
                return self.onUnresolvedUserClick(e);
            });
            this.parent.validateMultipleMatchError();
            this.parent.validateNoMatchError();
        },

        onUnresolvedUserClick: function (e) {
            e = Office.Controls.Utils.getEvent(e);
            this.parent.autofill.flushContent();
            this.parent.autofill.setServerEntries(this.optionsList);
            var self = this;
            this.parent.autofill.open(function (selectedPrincipal) {
                self.onAutofillClick(selectedPrincipal);
            });
            this.addKeyListenerForAutofill();
            this.parent.autofill.focusOnFirstElement();
            Office.Controls.Utils.cancelEvent(e);
            return false;
        },

        addKeyListenerForAutofill: function () {
            var autofillElementsLiTags = this.parent.autofill.root.querySelectorAll('li'), i;
            for (i = 0; i < autofillElementsLiTags.length; i++) {
                var li = autofillElementsLiTags[i];
                var self = this;
                Office.Controls.Utils.addEventListener(li, 'keydown', function (e) {
                    return self.onAutofillKeyDown(e);
                });
            }
        },

        onAutofillKeyDown: function (e) {
            var key = Office.Controls.Utils.getEvent(e),
            target = Office.Controls.Utils.getTarget(key);
            if (key.keyCode === 38) { // 'up arrow'
                if (target.previousSibling !== null) {
                    this.parent.autofill.changeFocus(target, target.previousSibling);
                    target.previousSibling.focus();
                } else if (target.parentNode.parentNode.previousSibling !== null) {
                    var autofillElementsUlTags = this.parent.root.querySelectorAll('ul'),
                    ul = autofillElementsUlTags[0];
                    this.parent.autofill.changeFocus(target, ul.lastChild);
                    ul.lastChild.focus();
                } else if (target.parentNode.parentNode.nextSibling !== null) {
                    var autofillElementsUlTags = this.parent.root.querySelectorAll('ul'),
                    ul = autofillElementsUlTags[1];
                    this.parent.autofill.changeFocus(target, ul.lastChild);
                    ul.lastChild.focus();
                } else {
                    var resultList = this.parent.root.querySelector('ul.ms-PeoplePicker-resultList');
                    this.parent.autofill.changeFocus(target, resultList.lastChild);
                    resultList.lastChild.focus();
                }
            } else if (key.keyCode === 40) { // 'down arrow'
                if (target.nextSibling !== null) {
                    this.parent.autofill.changeFocus(target, target.nextSibling);
                    target.nextSibling.focus();
                } else if (target.parentNode.parentNode.nextSibling !== null) {
                    var autofillElementsUlTags = this.parent.root.querySelectorAll('ul'),
                    ul = autofillElementsUlTags[1];
                    this.parent.autofill.changeFocus(target, ul.firstChild);
                    ul.firstChild.focus();
                } else if (target.parentNode.parentNode.previousSibling !== null) {
                    var autofillElementsUlTags = this.parent.root.querySelectorAll('ul'),
                    ul = autofillElementsUlTags[0];
                    this.parent.autofill.changeFocus(target, ul.firstChild);
                    ul.firstChild.focus();
                } else {
                    var resultList = this.parent.root.querySelector('ul.ms-PeoplePicker-resultList');
                    this.parent.autofill.changeFocus(target, resultList.firstChild);
                    resultList.firstChild.focus();
                }
            } else if (key.keyCode === 9 || key.keyCode === 13) { // 'tab' or 'enter'
                var personId = this.parent.autofill.getPersonIdFromListElement(target);
                this.onAutofillClick(this.parent.autofill.entries[personId]);
                Office.Controls.Utils.cancelEvent(e);
            }
            return true;
        },

        resolveTo: function (principal) {
            Office.Controls.PeoplePicker.copyToRecord(this.Record, principal);
            this.Record.text = principal.displayName;
            this.Record.isResolved = true;
            if (this.parent.enableCache) {
                this.parent.addToCache(principal);
            }
            Office.Controls.Utils.removeClass(this.Node, 'has-error');
            var primaryTextNode = this.Node.querySelector('div.ms-Persona-primaryText');
            primaryTextNode.innerHTML = Office.Controls.Utils.htmlEncode(principal.displayName);
            this.updateHoverText();
        },

        refresh: function (principal) {
            Office.Controls.PeoplePicker.copyToRecord(this.Record, principal);
            this.Record.text = principal.displayName;
            var primaryTextNode = this.Node.querySelector('div.ms-Persona-primaryText');
            primaryTextNode.innerHTML = Office.Controls.Utils.htmlEncode(principal.displayName);
        },

        unresolve: function () {
            this.Record.isResolved = false;
            var primaryTextNode = this.Node.querySelector('div.ms-Persona-primaryText');
            primaryTextNode.innerHTML = Office.Controls.Utils.htmlEncode(this.Record.text);
            this.updateHoverText();
        },

        updateHoverText: function () {
            var userLabel = this.Node.querySelector('div.ms-Persona-primaryText');
            userLabel.title = Office.Controls.Utils.htmlEncode(this.Record.text);
            this.Node.querySelector('div.ms-PeoplePicker-personaRemove').title = Office.Controls.Utils.formatString(Office.Controls.peoplePickerTemplates.getString(Office.Controls.Utils.htmlEncode('PP_RemovePerson')), this.Record.text);
        },

        onAutofillClick: function (selectedPrincipal) {
            this.parent.onRemove(this.parent, this.Record.principalInfo);
            this.resolveTo(selectedPrincipal);
            this.parent.refreshInputField();
            this.principalOptions = null;
            this.optionsList = null;
            if (this.parent.enableCache) {
                this.parent.addToCache(selectedPrincipal);
            }
            this.parent.validateMultipleMatchError();
            this.parent.autofill.close();
            this.parent.textInput.focus();
            this.parent.onAdd(this.parent, this.Record.principalInfo);
            this.parent.onChange(this.parent);
        }
    };

    Office.Controls.PeoplePicker.autofillContainer = function (parent) {
        this.entries = {};
        this.cachedEntries = new Array(0);
        this.serverEntries = new Array(0);
        this.parent = parent;
        this.root = parent.autofillElement;
        if (!Office.Controls.PeoplePicker.autofillContainer.boolBodyHandlerAdded) {
            Office.Controls.Utils.addEventListener(document.body, 'click', function (e) {
                return Office.Controls.PeoplePicker.autofillContainer.bodyOnClick(e);
            });
            Office.Controls.PeoplePicker.autofillContainer.boolBodyHandlerAdded = true;
        }
    };
    Office.Controls.PeoplePicker.autofillContainer.getControlRootFromSubElement = function (element) {
        while (element && element.nodeName.toLowerCase() !== 'body') {
            if (element.className.indexOf('office office-peoplepicker') !== -1) {
                return element;
            }
            element = element.parentNode;
        }
        return null;
    };
    Office.Controls.PeoplePicker.autofillContainer.bodyOnClick = function (e) {
        if (!Office.Controls.PeoplePicker.autofillContainer.currentOpened) {
            return true;
        }
        var click = Office.Controls.Utils.getEvent(e),
        target = Office.Controls.Utils.getTarget(click),
        controlRoot = Office.Controls.PeoplePicker.autofillContainer.getControlRootFromSubElement(target);
        if (!target || controlRoot !== Office.Controls.PeoplePicker.autofillContainer.currentOpened.parent.root) {
            Office.Controls.PeoplePicker.autofillContainer.currentOpened.close();
        }
        return true;
    };
    Office.Controls.PeoplePicker.autofillContainer.prototype = {
        _parent: null,
        _root: null,
        IsDisplayed: false,

        get_isDisplayed: function () {
            return this.IsDisplayed;
        },

        set_isDisplayed: function (value) {
            this.IsDisplayed = value;
            return value;
        },

        setCachedEntries: function (entries) {
            this.cachedEntries = entries;
            this.entries = {};
            var length = entries.length, i;
            for (i = 0; i < length; i++) {
                this.entries[entries[i].id] = entries[i];
            }
        },

        getCachedEntries: function () {
            return this.cachedEntries;
        },

        getServerEntries: function () {
            return this.serverEntries;
        },

        setServerEntries: function (entries) {
            if (this.parent.enableCache === true) {
                var newServerEntries = new Array(0),
                length = entries.length, i;
                for (i = 0; i < length; i++) {
                    var currentEntry = entries[i];
                    if (Office.Controls.Utils.isNullOrUndefined(this.entries[currentEntry.id])) {
                        this.entries[entries[i].id] = entries[i];
                        newServerEntries.push(currentEntry);
                    }
                }
                this.serverEntries = newServerEntries;
            } else {
                this.entries = {};
                var length = entries.length, i;
                for (i = 0; i < length; i++) {
                    this.entries[entries[i].id] = entries[i];
                }
                this.serverEntries = entries;
            }
        },

        renderList: function (handler) {
            this.root.innerHTML = Office.Controls.peoplePickerTemplates.generateAutofillListTemplate(this.cachedEntries, this.serverEntries, this.parent.numberOfResults, this.parent.showImage);
            var isTabKey = false,
            autofillElementsLinkTags = this.root.querySelectorAll('a'),
            self = this, i;
            for (i = 0; i < autofillElementsLinkTags.length; i++) {
                var link = autofillElementsLinkTags[i];
                Office.Controls.Utils.addEventListener(link, 'click', function (e) {
                    return self.onEntryClick(e, handler);
                });
                Office.Controls.Utils.addEventListener(link, 'focus', function (e) {
                    return self.onEntryFocus(e);
                });
                Office.Controls.Utils.addEventListener(link, 'blur', function (e) {
                    return self.onEntryBlur(e, isTabKey);
                });
            }
            var autofillElementsLiTags = this.root.querySelectorAll('li');
            if (autofillElementsLiTags.length > 0) {
                Office.Controls.Utils.addClass(autofillElementsLiTags[0], 'ms-PeoplePicker-resultAddedForSelect');
            }
            if (this.parent.showImage) {
                for (i = 0; i < autofillElementsLiTags.length; i++) {
                    var li = autofillElementsLiTags[i];
                    var image = li.querySelector('.ms-PeoplePicker-Persona-image');
                    var personId = this.getPersonIdFromListElement(li);
                    (function (self, image, personId) {
                        self.parent.dataProvider.getImageAsync(personId, function (error, imgSrc) {
                            if (imgSrc != null) {
                                image.style.backgroundImage = "url('" + imgSrc + "')";
                                if (!Office.Controls.Utils.isNullOrUndefined(self.entries[personId])) {
                                    self.entries[personId].imgSrc = imgSrc;
                                }
                            }
                        });
                    })(this, image, personId);
                }
            }
        },

        flushContent: function () {
            var entry = this.root.querySelectorAll('div.ms-PeoplePicker-resultGroups'), i;
            for (i = 0; i < entry.length; i++) {
                this.root.removeChild(entry[i]);
            }
            this.entries = {};
            this.serverEntries = new Array(0);
            this.cachedEntries = new Array(0);
        },

        open: function (handler) {
            this.renderList(handler);
            this.IsDisplayed = true;
            Office.Controls.PeoplePicker.autofillContainer.currentOpened = this;
            if (!Office.Controls.Utils.containClass(this.parent.actualRoot, 'is-active')) {
                Office.Controls.Utils.addClass(this.parent.actualRoot, 'is-active');
            }
        },

        close: function () {
            this.IsDisplayed = false;
            Office.Controls.Utils.removeClass(this.parent.actualRoot, 'is-active');
        },

        openSearchingLoadingStatus: function (searchingName, isAppended) {
            if (isAppended == false) {
                this.root.innerHTML = Office.Controls.peoplePickerTemplates.generateSerachingLoadingTemplate();
            }
            else {
                var resultNode = this.root.querySelector('div.ms-PeoplePicker-resultGroups')
                if (resultNode !=null){
                    resultNode.innerHTML += Office.Controls.peoplePickerTemplates.generateSerachingLoadingTemplate();
                }
                else {
                    this.root.innerHTML = Office.Controls.peoplePickerTemplates.generateSerachingLoadingTemplate();
                }
            }
            this.IsDisplayed = true;
            Office.Controls.PeoplePicker.autofillContainer.currentOpened = this;
            if (!Office.Controls.Utils.containClass(this.parent.actualRoot, 'is-active')) {
                Office.Controls.Utils.addClass(this.parent.actualRoot, 'is-active');
            }
        },

        closeSearchingLoadingStatus: function () {
            this.IsDisplayed = false;
            Office.Controls.Utils.removeClass(this.parent.actualRoot, 'is-active');
        },

        onEntryClick: function (e, handler) {
            var click = Office.Controls.Utils.getEvent(e),
            target = Office.Controls.Utils.getTarget(click),
            listItem = this.getParentListItem(target),
            personId = this.getPersonIdFromListElement(listItem);
            handler(this.entries[personId]);
            this.flushContent();
            return true;
        },

        focusOnFirstElement: function () {
            var first = this.root.querySelector('li.ms-PeoplePicker-result');
            if (!Office.Controls.Utils.isNullOrUndefined(first)) {
                first.focus();
            }
        },

        onKeyDownFromInput: function (key) {
            var target = this.root.querySelector("li.ms-PeoplePicker-resultAddedForSelect");
            if (key.keyCode === 38) { // 'up arrow'
                if (target.previousSibling !== null) {
                    this.changeFocus(target, target.previousSibling);
                    target.previousSibling.focus();
                } else if (target.parentNode.parentNode.previousSibling !== null) {
                    var autofillElementsUlTags = this.parent.root.querySelectorAll('ul'),
                    ul = autofillElementsUlTags[0];
                    this.parent.autofill.changeFocus(target, ul.lastChild);
                    ul.lastChild.focus();
                } else if (target.parentNode.parentNode.nextSibling !== null) {
                    var autofillElementsUlTags = this.parent.root.querySelectorAll('ul'),
                    ul = autofillElementsUlTags[1];
                    this.parent.autofill.changeFocus(target, ul.lastChild);
                    ul.lastChild.focus();
                } else {
                    var resultList = this.root.querySelector('ul.ms-PeoplePicker-resultList');
                    this.changeFocus(target, resultList.lastChild);
                    resultList.lastChild.focus();
                }
            } else if (key.keyCode === 40) { // 'down arrow'
                if (target.nextSibling !== null) {
                    this.changeFocus(target, target.nextSibling);
                    target.nextSibling.focus();
                } else if (target.parentNode.parentNode.nextSibling !== null) {
                    var autofillElementsUlTags = this.parent.root.querySelectorAll('ul'),
                    ul = autofillElementsUlTags[1];
                    this.parent.autofill.changeFocus(target, ul.firstChild);
                    ul.firstChild.focus();
                } else if (target.parentNode.parentNode.previousSibling !== null) {
                    var autofillElementsUlTags = this.parent.root.querySelectorAll('ul'),
                    ul = autofillElementsUlTags[0];
                    this.parent.autofill.changeFocus(target, ul.firstChild);
                    ul.firstChild.focus();
                } else {
                    var resultList = this.root.querySelector('ul.ms-PeoplePicker-resultList');
                    this.changeFocus(target, resultList.firstChild);
                    resultList.firstChild.focus();
                }
            }
            this.parent.textInput.focus();
            return true;
        },

        changeFocus: function (lastElement, nextElement) {
            Office.Controls.Utils.removeClass(lastElement, 'ms-PeoplePicker-resultAddedForSelect');
            Office.Controls.Utils.addClass(nextElement, 'ms-PeoplePicker-resultAddedForSelect');
        },

        getPersonIdFromListElement: function (listElement) {
            return listElement.attributes.getNamedItem('data-office-peoplepicker-value').value;
        },

        getParentListItem: function (element) {
            while (element && element.nodeName.toLowerCase() !== 'li') {
                element = element.parentNode;
            }
            return element;
        },

        onEntryFocus: function (e) {
            var target = Office.Controls.Utils.getTarget(e);
            target = this.getParentListItem(target);
            if (!Office.Controls.Utils.containClass(target, 'office-peoplepicker-autofill-focus')) {
                Office.Controls.Utils.addClass(target, 'office-peoplepicker-autofill-focus');
            }
            return false;
        },

        onEntryBlur: function (e, isTabKey) {
            var target = Office.Controls.Utils.getTarget(e);
            target = this.getParentListItem(target);
            Office.Controls.Utils.removeClass(target, 'office-peoplepicker-autofill-focus');
            if (isTabKey) {
                var next = target.nextSibling;
                if (next && (next.nextSibling.className.toLowerCase() === 'ms-PeoplePicker-searchMore js-searchMore'.toLowerCase())) {
                    Office.Controls.PeoplePicker.autofillContainer.currentOpened.close();
                }
            }
            return false;
        }
    };

    Office.Controls.PeoplePicker.Parameters = function () { };

    Office.Controls.PeoplePicker.cancelToken = function () {
        this.IsCanceled = false;
    };
    Office.Controls.PeoplePicker.cancelToken.prototype = {
        IsCanceled: false,

        get_isCanceled: function () {
            return this.IsCanceled;
        },

        set_isCanceled: function (value) {
            this.IsCanceled = value;
            return value;
        },

        cancel: function () {
            this.IsCanceled = true;
        }
    };

    Office.Controls.PeoplePicker.ValidationError = function () {};
    Office.Controls.PeoplePicker.ValidationError.createMultipleMatchError = function () {
        var err = new Office.Controls.PeoplePicker.ValidationError();
        err.errorName = 'MultipleMatch';
        err.localizedErrorMessage = Office.Controls.peoplePickerTemplates.getString('PP_MultipleMatch');
        return err;
    };
    Office.Controls.PeoplePicker.ValidationError.createNoMatchError = function () {
        var err = new Office.Controls.PeoplePicker.ValidationError();
        err.errorName = 'NoMatch';
        err.localizedErrorMessage = Office.Controls.peoplePickerTemplates.getString('PP_NoMatch');
        return err;
    };
    Office.Controls.PeoplePicker.ValidationError.createServerProblemError = function () {
        var err = new Office.Controls.PeoplePicker.ValidationError();
        err.errorName = 'ServerProblem';
        err.localizedErrorMessage = Office.Controls.peoplePickerTemplates.getString('PP_ServerProblem');
        return err;
    };
    Office.Controls.PeoplePicker.ValidationError.prototype = {
        errorName: null,
        localizedErrorMessage: null
    };

    Office.Controls.PeoplePicker.mruCache = function () {
        this.isCacheAvailable = this.checkCacheAvailability();
        if (!this.isCacheAvailable) {
            return;
        }
        this.initializeCache();
    };

    Office.Controls.PeoplePicker.mruCache.getInstance = function () {
        if (!Office.Controls.PeoplePicker.mruCache.instance) {
            Office.Controls.PeoplePicker.mruCache.instance = new Office.Controls.PeoplePicker.mruCache();
        }
        return Office.Controls.PeoplePicker.mruCache.instance;
    };

    Office.Controls.PeoplePicker.mruCache.prototype = {
        isCacheAvailable: false,
        _localStorage: null,
        _dataObject: null,

        get: function (key, maxResults) {
            if (Office.Controls.Utils.isNullOrUndefined(maxResults) || !maxResults) {
                maxResults = 2147483647;
            }
            var numberOfResults = 0,
            results = new Array(0),
            cache = this.dataObject.cacheMapping[Office.Controls.Runtime.context.HostUrl],
            cacheLength = cache.length, i;
            for (i = cacheLength; i > 0 && numberOfResults < maxResults; i--) {
                var candidate = cache[i - 1];
                if (this.entityMatches(candidate, key)) {
                    results.push(candidate);
                    numberOfResults += 1;
                }
            }
            return results;
        },

        set: function (entry) {
            var cache = this.dataObject.cacheMapping[Office.Controls.Runtime.context.HostUrl],
            cacheSize = cache.length,
            alreadyThere = false, i;
            for (i = 0; i < cacheSize; i++) {
                var cacheEntry = cache[i];
                if (cacheEntry.id === entry.id) {
                    cache.splice(i, 1);
                    alreadyThere = true;
                    break;
                }
            }
            if (!alreadyThere) {
                if (cacheSize >= 200) {
                    cache.splice(0, 1);
                }
            }
            cache.push(entry);
            this.cacheWrite('Office.PeoplePicker.Cache', Office.Controls.Utils.serializeJSON(this.dataObject));
        },

        entityMatches: function (candidate, key) {
            if (Office.Controls.Utils.isNullOrEmptyString(key) || Office.Controls.Utils.isNullOrUndefined(candidate)) {
                return false;
            }
            key = key.toLowerCase();
            var userNameKey = candidate.id;
            if (Office.Controls.Utils.isNullOrUndefined(userNameKey)) {
                userNameKey = '';
            }
            var divideIndex = userNameKey.indexOf('\\');
            if (divideIndex !== -1 && divideIndex !== userNameKey.length - 1) {
                userNameKey = userNameKey.substr(divideIndex + 1);
            }
            var emailKey = candidate.Email;
            if (Office.Controls.Utils.isNullOrUndefined(emailKey)) {
                emailKey = '';
            }
            var atSignIndex = emailKey.indexOf('@');
            if (atSignIndex !== -1) {
                emailKey = emailKey.substr(0, atSignIndex);
            }
            if (Office.Controls.Utils.isNullOrUndefined(candidate.displayName)) {
                candidate.displayName = '';
            }
            if ((!userNameKey.toLowerCase().indexOf(key)) || (!emailKey.toLowerCase().indexOf(key)) || (!candidate.displayName.toLowerCase().indexOf(key))) {
                return true;
            }
            return false;
        },

        initializeCache: function () {
            var cacheData = this.cacheRetreive('Office.PeoplePicker.Cache');
            if (Office.Controls.Utils.isNullOrEmptyString(cacheData)) {
                this.dataObject = new Office.Controls.PeoplePicker.mruCache.mruData();
            } else {
                var datas = Office.Controls.Utils.deserializeJSON(cacheData);
                if (datas.cacheVersion) {
                    this.dataObject = new Office.Controls.PeoplePicker.mruCache.mruData();
                    this.cacheDelete('Office.PeoplePicker.Cache');
                } else {
                    this.dataObject = datas;
                }
            }
            if (Office.Controls.Utils.isNullOrUndefined(this.dataObject.cacheMapping[Office.Controls.Runtime.context.HostUrl])) {
                this.dataObject.cacheMapping[Office.Controls.Runtime.context.HostUrl] = new Array(0);
            }
        },

        checkCacheAvailability: function () {
            try {
                if (typeof window.self.localStorage === 'undefined') {
                    return false;
                }
                this.localStorage = window.self.localStorage;
                return true;
            } catch (e) {
                return false;
            }
        },

        cacheRetreive: function (key) {
            return this.localStorage.getItem(key);
        },

        cacheWrite: function (key, value) {
            this.localStorage.setItem(key, value);
        },

        cacheDelete: function (key) {
            this.localStorage.removeItem(key);
        }
    };

    Office.Controls.PeoplePicker.mruCache.mruData = function () {
        this.cacheMapping = {};
        this.cacheVersion = 0;
    };

    Office.Controls.PeoplePickerResourcesDefaults = function () {
    };

    Office.Controls.peoplePickerTemplates = function () {
    };
    Office.Controls.peoplePickerTemplates.getString = function (stringName) {
        var newName = 'PeoplePicker' + stringName.substr(3);
        if (Office.Controls.PeoplePicker.resourceStrings.hasOwnProperty(newName)) {
            return Office.Controls.PeoplePicker.resourceStrings[newName];
        }
        return Office.Controls.Utils.getStringFromResource('PeoplePicker', stringName);
    };

    Office.Controls.peoplePickerTemplates.getDefaultText = function (allowMultipleSelections) {
        if (allowMultipleSelections) {
            return Office.Controls.peoplePickerTemplates.getString('PP_DefaultMessagePlural');
        }
        return Office.Controls.peoplePickerTemplates.getString('PP_DefaultMessage');
    };

    Office.Controls.peoplePickerTemplates.generateControlTemplate = function (inputName, allowMultipleSelections, inputHint) {
        var defaultText;
        if (Office.Controls.Utils.isNullOrEmptyString(inputHint)) {
            defaultText = Office.Controls.Utils.htmlEncode(Office.Controls.peoplePickerTemplates.getDefaultText(allowMultipleSelections));
        } else {
            defaultText = Office.Controls.Utils.htmlEncode(inputHint);
        }
        var body = '<div class=\"ms-PeoplePicker\">';
        body += '<input type=\"hidden\"';
        if (!Office.Controls.Utils.isNullOrEmptyString(inputName)) {
            body += ' name=\"' + Office.Controls.Utils.htmlEncode(inputName) + '\"';
        }
        body += '/>';
        body += '<div class=\"ms-PeoplePicker-searchBox ms-PeoplePicker-searchBoxAdded\">';
        body += '<span class=\"office-peoplepicker-default\">' + defaultText + '</span>';
        body += '<div class=\"office-peoplepicker-recordList\"></div>';
        body += '<input type=\"text\" class=\"ms-PeoplePicker-searchField\" size=\"1\" autocorrect=\"off\" autocomplete=\"off\" autocapitalize=\"off\" />';
        body += '</div>';
        body += '<div class=\"ms-PeoplePicker-results\">';
        body += '</div>';
        body += Office.Controls.peoplePickerTemplates.generateAlertNode();
        body += '</div>';
        return body;
    };

    Office.Controls.peoplePickerTemplates.generateErrorTemplate = function (ErrorMessage) {
        var innerHtml = '<span class=\"office-peoplepicker-error\">';
        innerHtml += Office.Controls.Utils.htmlEncode(ErrorMessage);
        innerHtml += '</span>';
        return innerHtml;
    };

    Office.Controls.peoplePickerTemplates.generateAutofillListItemTemplate = function (principal, isCached, showImage) {
        var titleText = Office.Controls.Utils.htmlEncode((Office.Controls.Utils.isNullOrEmptyString(principal.Email)) ? '' : principal.Email),
        itemHtml = '<li tabindex=\"0\" class=\"ms-PeoplePicker-result\" data-office-peoplepicker-value=\"' + Office.Controls.Utils.htmlEncode(principal.id) + '\" title=\"' + titleText + '\">';
        itemHtml += '<a onclick=\"return false;\" href=\"#\" tabindex=\"-1\">';
        itemHtml += '<div class=\"ms-Persona ms-PersonaAdded\">';
        if (showImage) {
            if (isCached && !Office.Controls.Utils.isNullOrUndefined(principal.imgSrc)) {
                if (Office.Controls.Utils.isIE10()) {
                    itemHtml += '<div class=\"ms-PeoplePicker-Persona-image\" style=\"display:block;background-image:url(\'' + principal.imgSrc + '\')\"></div>';
                } else {
                    itemHtml += '<img class=\"ms-PeoplePicker-Persona-image\" style=\"background-image:url(\'' + principal.imgSrc + '\')\">';
                }
            } else {
                if (Office.Controls.Utils.isIE10()) {
                    itemHtml += '<div class=\"ms-PeoplePicker-Persona-image\" style=\"display:block\"></div>';
                } else {
                    itemHtml += '<img class=\"ms-PeoplePicker-Persona-image\">';
                }
            }
        } else {
            itemHtml += '<div class=\"ms-Persona-image-placeholder\"></div>';
        }
        if (Office.Controls.Utils.isFirefox()) {
            itemHtml += '<div class=\"ms-Persona-details\" style=\"max-width:100%; width: auto;\">';
        } else {
            itemHtml += '<div class=\"ms-Persona-details\">';
        }
        itemHtml += '<div class=\"ms-Persona-primaryText\" >' + Office.Controls.Utils.htmlEncode(principal.displayName) + '</div>';
        if (!Office.Controls.Utils.isNullOrEmptyString(principal.description)) {
            itemHtml += '<div class=\"ms-Persona-secondaryText\" >' + Office.Controls.Utils.htmlEncode(principal.description) + '</div>';
        }
        itemHtml += '</div></div></a></li>';
        return itemHtml;
    };

    Office.Controls.peoplePickerTemplates.generateAutofillListTemplate = function (cachedEntries, serverEntries, maxCount, showImage) {
        var html = '<div class=\"ms-PeoplePicker-resultGroups\">',
        actualCount = cachedEntries.length + serverEntries.length;
        if (Office.Controls.Utils.isNullOrUndefined(cachedEntries)) {
            cachedEntries = new Array(0);
        }
        if (Office.Controls.Utils.isNullOrUndefined(serverEntries) || cachedEntries.length >= maxCount) {
            serverEntries = new Array(0);
        }
        if (actualCount > maxCount && cachedEntries.length < maxCount) {
            serverEntries = serverEntries.slice(0, maxCount - cachedEntries.length);
        }
        html += Office.Controls.peoplePickerTemplates.generateAutofillGroupTemplate(cachedEntries, true, showImage, true);
        html += Office.Controls.peoplePickerTemplates.generateAutofillGroupTemplate(serverEntries, false, showImage, (cachedEntries.length > 0));
        html += '</div>';
        html += Office.Controls.peoplePickerTemplates.generateAutofillFooterTemplate(actualCount, maxCount);
        return html;
    };

    Office.Controls.peoplePickerTemplates.generateAutofillGroupTemplate = function (principals, isCached, showImage, showTitle) {
        var listHtml = '',
        cachedGrouptTitle = Office.Controls.Utils.htmlEncode(Office.Controls.peoplePickerTemplates.getString('PP_SearchResultRecentGroup')),
        searchedGroupTitle = Office.Controls.Utils.htmlEncode(Office.Controls.peoplePickerTemplates.getString('PP_SearchResultMoreGroup')), i,
        groupTitle = isCached ? cachedGrouptTitle : searchedGroupTitle;
        if (!principals.length) {
            return listHtml;
        }
        listHtml += '<div class=\"ms-PeoplePicker-resultGroup\">';
        if (showTitle) {
            listHtml += '<div class=\"ms-PeoplePicker-resultGroupTitle\">' + groupTitle + '</div>';
        }
        listHtml += '<ul class=\"ms-PeoplePicker-resultList\" id=\"' + groupTitle + '\">';
        for (i = 0; i < principals.length; i++) {
            listHtml += Office.Controls.peoplePickerTemplates.generateAutofillListItemTemplate(principals[i], isCached, showImage);
        }
        listHtml += '</ul></div>';
        return listHtml;
    };

    Office.Controls.peoplePickerTemplates.generateAutofillFooterTemplate = function (count, maxCount) {
        var footerHtml = '<div class=\"ms-PeoplePicker-searchMore js-searchMore\">';
        footerHtml += '<div class=\"ms-PeoplePicker-searchMoreIcon\"></div>';
        var footerText;
        if (count >= maxCount) {
            footerText = Office.Controls.Utils.formatString(Office.Controls.peoplePickerTemplates.getString('PP_ShowingTopNumberOfResults'), maxCount.toString());
        } else if (count > 1) {
            footerText = Office.Controls.Utils.formatString(Office.Controls.peoplePickerTemplates.getString('PP_MultipleResults'), count.toString());
        } else if (count > 0) {
            footerText = Office.Controls.Utils.formatString(Office.Controls.peoplePickerTemplates.getString('PP_SingleResult'), count.toString());
        } else {
            footerText = Office.Controls.peoplePickerTemplates.getString('PP_NoResult');
        }
        footerText = Office.Controls.Utils.htmlEncode(footerText);
        footerHtml += '<div class=\"ms-PeoplePicker-searchMorePrimary ms-PeoplePicker-searchMorePrimaryAdded\">' + footerText + '</div>';
        footerHtml += '</div>';
        return footerHtml;
    };

    Office.Controls.peoplePickerTemplates.generateSerachingLoadingTemplate = function () {
        var searchingLable = Office.Controls.Utils.htmlEncode(Office.Controls.peoplePickerTemplates.getString('PP_Searching'));
        var searchingLoadingHtml = '<div class=\"ms-PeoplePicker-searchMore js-searchMore is-searching\">';
        searchingLoadingHtml += '<div class=\"ms-PeoplePicker-searchMoreIconFixed\"></div>';
        searchingLoadingHtml += '<div class=\"ms-PeoplePicker-searchMorePrimary ms-PeoplePicker-searchMorePrimaryAdded\">' + searchingLable + '</div>';
        searchingLoadingHtml += '</div>';
        return searchingLoadingHtml;
    };

    Office.Controls.peoplePickerTemplates.generateRecordTemplate = function (record, allowMultipleSelections, showImage) {
        var recordHtml;
        var userRecordClass = 'ms-PeoplePicker-persona';
        if (!allowMultipleSelections) {
            userRecordClass += ' ms-PeoplePicker-personaForSingleAdded';
        }
        if (record.isResolved) {
            recordHtml = '<div class=\"' + userRecordClass + '\" tabindex=\"0\">';
        } else {
            recordHtml = '<div class=\"' + userRecordClass + ' ' + 'has-error' + '\" tabindex=\"0\">';
        }
        recordHtml += '<div class=\"ms-Persona ms-Persona--xs\" >';
        if (showImage) {
            if (Office.Controls.Utils.isNullOrUndefined(record.imgSrc)) {
                if (Office.Controls.Utils.isIE10()) {
                    recordHtml += '<div class=\"ms-PeoplePicker-Persona-image\" style=\"display:block\"></div>';
                } else {
                    recordHtml += '<img class=\"ms-PeoplePicker-Persona-image\">';
                }
            } else  {
                if (Office.Controls.Utils.isIE10()) {
                    recordHtml += '<div class=\"ms-PeoplePicker-Persona-image\" style=\"display:block;background-image:url(\'' + record.imgSrc + '\')\"></div>';
                } else {
                    recordHtml += '<img class=\"ms-PeoplePicker-Persona-image\" style=\"background-image:url(\'' + record.imgSrc + '\')\">';
                }
            }
        }
        recordHtml += '<div class=\"ms-Persona-details\">';
        recordHtml += '<div class=\"ms-Persona-primaryText ms-Persona-primaryTextForResolvedUserAdded\">' + Office.Controls.Utils.htmlEncode(record.text);
        recordHtml += '</div></div></div>';
        if (showImage) {
            recordHtml += '<div class=\"ms-PeoplePicker-personaRemove\">';
        } else {
            recordHtml += '<div class=\"ms-PeoplePicker-personaRemove\">';
        }
        recordHtml += '<i class=\"ms-Icon ms-Icon--x ms-Icon-added\">';
        recordHtml += '</i></div>';
        recordHtml += '</div>';
        return recordHtml;
    };

    Office.Controls.peoplePickerTemplates.generateAlertNode = function () {
        var alertHtml = '<div role=\"alert\" class=\"office-peoplepicker-alert\">';
        alertHtml += '</div>';
        return alertHtml;
    };

    if (Office.Controls.PrincipalInfo.registerClass) { Office.Controls.PrincipalInfo.registerClass('Office.Controls.PrincipalInfo'); }
    if (Office.Controls.PeoplePickerRecord.registerClass) { Office.Controls.PeoplePickerRecord.registerClass('Office.Controls.PeoplePickerRecord'); }
    if (Office.Controls.PeoplePicker.registerClass) { Office.Controls.PeoplePicker.registerClass('Office.Controls.PeoplePicker'); }
    if (Office.Controls.PeoplePicker.internalPeoplePickerRecord.registerClass) { Office.Controls.PeoplePicker.internalPeoplePickerRecord.registerClass('Office.Controls.PeoplePicker.internalPeoplePickerRecord'); }
    if (Office.Controls.PeoplePicker.autofillContainer.registerClass) { Office.Controls.PeoplePicker.autofillContainer.registerClass('Office.Controls.PeoplePicker.autofillContainer'); }
    if (Office.Controls.PeoplePicker.Parameters.registerClass) { Office.Controls.PeoplePicker.Parameters.registerClass('Office.Controls.PeoplePicker.Parameters'); }
    if (Office.Controls.PeoplePicker.cancelToken.registerClass) { Office.Controls.PeoplePicker.cancelToken.registerClass('Office.Controls.PeoplePicker.cancelToken');}
    if (Office.Controls.PeoplePicker.ValidationError.registerClass) { Office.Controls.PeoplePicker.ValidationError.registerClass('Office.Controls.PeoplePicker.ValidationError'); }
    if (Office.Controls.PeoplePicker.mruCache.registerClass) { Office.Controls.PeoplePicker.mruCache.registerClass('Office.Controls.PeoplePicker.mruCache');}
    if (Office.Controls.PeoplePicker.mruCache.mruData.registerClass) { Office.Controls.PeoplePicker.mruCache.mruData.registerClass('Office.Controls.PeoplePicker.mruCache.mruData'); }
    if (Office.Controls.PeoplePickerResourcesDefaults.registerClass) { Office.Controls.PeoplePickerResourcesDefaults.registerClass('Office.Controls.PeoplePickerResourcesDefaults'); }
    if (Office.Controls.peoplePickerTemplates.registerClass) { Office.Controls.peoplePickerTemplates.registerClass('Office.Controls.peoplePickerTemplates'); }
    Office.Controls.PeoplePicker.resourceStrings = {};
    Office.Controls.PeoplePicker.autofillContainer.currentOpened = null;
    Office.Controls.PeoplePicker.autofillContainer.boolBodyHandlerAdded = false;
    Office.Controls.PeoplePicker.ValidationError.multipleMatchName = 'MultipleMatch';
    Office.Controls.PeoplePicker.ValidationError.noMatchName = 'NoMatch';
    Office.Controls.PeoplePicker.ValidationError.serverProblemName = 'ServerProblem';
    Office.Controls.PeoplePicker.mruCache.instance = null;
    Office.Controls.PeoplePickerResourcesDefaults.PP_NoMatch = 'We couldn\'t find an exact match.';
    Office.Controls.PeoplePickerResourcesDefaults.PP_ServerProblem = 'Sorry, we\'re having trouble reaching the server.';
    Office.Controls.PeoplePickerResourcesDefaults.PP_DefaultMessagePlural = 'Enter names or email addresses...';
    Office.Controls.PeoplePickerResourcesDefaults.PP_MultipleMatch = 'Multiple entries matched, please click to resolve.';
    Office.Controls.PeoplePickerResourcesDefaults.PP_NoResult = 'No results found';
    Office.Controls.PeoplePickerResourcesDefaults.PP_SingleResult = 'Showing {0} result';
    Office.Controls.PeoplePickerResourcesDefaults.PP_MultipleResults = 'Showing {0} results';
    Office.Controls.PeoplePickerResourcesDefaults.PP_ShowingTopNumberOfResults = 'Showing top {0} results';
    Office.Controls.PeoplePickerResourcesDefaults.PP_Searching = 'Searching...';
    Office.Controls.PeoplePickerResourcesDefaults.PP_RemovePerson = 'Remove person or group {0}';
    Office.Controls.PeoplePickerResourcesDefaults.PP_DefaultMessage = 'Enter a name or email address...';
    Office.Controls.PeoplePickerResourcesDefaults.PP_SearchResultRecentGroup = 'Recent';
    Office.Controls.PeoplePickerResourcesDefaults.PP_SearchResultMoreGroup = 'More';
})();

