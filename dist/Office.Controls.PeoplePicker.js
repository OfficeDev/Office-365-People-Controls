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
        Description: null,
        PersonId: null,
        imgSrc: null,
        principalInfo: null
    }

    Office.Controls.PeoplePicker = function (root, dataProvider, parameterObject) {
        try {
            if (typeof root !== 'object' || typeof dataProvider !== 'object' || (!Office.Controls.Utils.isNullOrUndefined(parameterObject) && typeof parameterObject !== 'object')) {
                Office.Controls.Utils.errorConsole('Invalid parameters type');
                return;
            }
            this.root = root;
            this.dataProvider = dataProvider;
            if (!Office.Controls.Utils.isNullOrUndefined(parameterObject)) {
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.allowMultipleSelections)) {
                    this.allowMultiple = (String(parameterObject.allowMultipleSelections) === "true");
                }
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.startSearchCharLength) && parameterObject.startSearchCharLength >= 1) {
                    this.startSearchCharLength = parameterObject.startSearchCharLength;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.delaySearchInterval)) {
                    this.delaySearchInterval = parameterObject.delaySearchInterval;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.enableCache)) {
                    this.enableCache = (String(parameterObject.enableCache) === "true");
                }
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.numberOfResults)) {
                    this.numberOfResults = parameterObject.numberOfResults;
                }
                if (!Office.Controls.Utils.isNullOrEmptyString(parameterObject.inputHint)) {
                    this.inputHint = parameterObject.inputHint;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.showValidationErrors)) {
                    this.showValidationErrors = (String(parameterObject.showValidationErrors) === "true");
                }
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.showImage)) {
                    this.showImage = (String(parameterObject.showImage) === "true");
                }
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.onAdded)) {
                    this.onAdded = parameterObject.onAdded;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.onRemoved)) {
                    this.onRemoved = parameterObject.onRemoved;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.onChange)) {
                    this.onChange = parameterObject.onChange;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.onFocus)) {
                    this.onFocus = parameterObject.onFocus;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.onBlur)) {
                    this.onBlur = parameterObject.onBlur;
                }
                this.onError = parameterObject.onError;

                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.resourceStrings)) {
                    Office.Controls.PeoplePicker.res = parameterObject.resourceStrings;
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
        record.DisplayName = info.DisplayName;
        record.Description = info.Description;
        record.PersonId = info.PersonId;
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

    Office.Controls.PeoplePicker.create = function (root, parameterObject) {
        return new Office.Controls.PeoplePicker(root, parameterObject);
    };
    Office.Controls.PeoplePicker.prototype = {
        allowMultiple: false,
        startSearchCharLength: 1,
        delaySearchInterval: 300,
        enableCache: true,
        inputHint: null,
        numberOfResults: 30,
        onAdded: Office.Controls.PeoplePicker.nopAddRemove,
        onRemoved: Office.Controls.PeoplePicker.nopAddRemove,
        onChange: Office.Controls.PeoplePicker.nopOperation,
        onFocus: Office.Controls.PeoplePicker.nopOperation,
        onBlur: Office.Controls.PeoplePicker.nopOperation,
        onError: null,
        dataProvider: null,
        showValidationErrors: true,
        showImage: false,
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
                    this.onRemoved(this, recordToRemove.principalInfo);
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
                record.text = p1.DisplayName;
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
            this.dataProvider.getPrincipals(userEmail, function (error, principalsReceived) {
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
            this.root.innerHTML = Office.Controls.peoplePickerTemplates.generateControlTemplate(inputName, this.allowMultiple, this.inputHint);
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
            if (keyEvent.keyCode === 27) {
                this.autofill.close();
            } else if (keyEvent.keyCode === 9 && this.autofill.IsDisplayed) {
                var focusElement = this.autofillElement.querySelector("li.ms-PeoplePicker-resultAddedForSelect");
                if (focusElement !== null) {
                    var personId = this.autofill.getPersonIdFromListElement(focusElement);
                    this.addResolvedPrincipal(this.autofill.entries[personId]);
                    this.autofill.flushContent();
                    Office.Controls.Utils.cancelEvent(e);
                    return false;
                }
                this.autofill.close();
            } else if ((keyEvent.keyCode === 40 || keyEvent.keyCode === 38) && this.autofill.IsDisplayed) {
                this.autofill.onKeyDownFromInput(keyEvent);
                keyEvent.preventDefault();
                keyEvent.stopPropagation();
                Office.Controls.Utils.cancelEvent(e);
                return false;
            } else if (keyEvent.keyCode === 37 && this.internalSelectedItems.length) {
                this.resolvedListRoot.lastChild.focus();
                Office.Controls.Utils.cancelEvent(e);
                return false;
            } else if (keyEvent.keyCode === 8) {
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
                keyEvent.preventDefault();
                keyEvent.stopPropagation();
                this.cancelLastRequest();
                this.attemptResolveInput();
                Office.Controls.Utils.cancelEvent(e);
                return false;
            } else if ((keyEvent.keyCode === 86 && keyEvent.ctrlKey) || (keyEvent.keyCode === 186)) {
                this.cancelLastRequest();
                window.setTimeout(function () {
                    self.textInput.value = Office.Controls.PeoplePicker.parseUserPaste(self.textInput.value);
                    self.attemptResolveInput();
                }, 0);
                return true;
            } else if (keyEvent.keyCode === 13 && keyEvent.shiftKey) {
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
                        self.displayLoadingIcon(currentValue);
                        self.removeValidationError('ServerProblem');
                        var token = new Office.Controls.PeoplePicker.cancelToken();
                        self.currentToken = token;
                        self.dataProvider.getPrincipals(self.textInput.value, function (error, principalsReceived) {
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
                    }
                    if (self.enableCache) {
                        self.displayCachedEntries();
                    }
                }
            }, self.delaySearchInterval);
        },

        onDataFetchError: function (message) {
            this.hideLoadingIcon();
            this.addValidationError(Office.Controls.PeoplePicker.ValidationError.createServerProblemError());
        },

        onDataReceived: function (principalsReceived) {
            this.currentPrincipalsChoices = {};
            var self = this, i;
            for (i = 0; i < principalsReceived.length; i++) {
                var principal = principalsReceived[i];
                this.currentPrincipalsChoices[principal.PersonId] = principal;
            }
            this.autofill.setServerEntries(principalsReceived);
            this.hideLoadingIcon();
            this.autofill.open(function (selectedPrincipal) {
                self.addResolvedPrincipal(selectedPrincipal);
            });
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

            if (this.startSearchCharLength === 0 && (this.allowMultiple === true || this.internalSelectedItems.length === 0)) {
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
            this.onRemoved(this, selectedPrincipal.principalInfo);
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
            if ((!this.allowMultiple) && (this.internalSelectedItems.length === 1)) {
                this.textInput.className = 'ms-PeoplePicker-searchFieldAddedForSingleSelectionHidden';
                this.textInput.setAttribute('readonly', 'readonly');
            } else {
                this.textInput.removeAttribute('readonly');
                this.textInput.className = 'ms-PeoplePicker-searchField ms-PeoplePicker-searchFieldAdded';
            }
        },

        displayLoadingIcon: function (searchingName) {
            this.autofill.openSearchingLoadingStatus(searchingName);
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
            record.text = principal.DisplayName;
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
            this.onAdded(this, record.principalInfo);
            this.onChange(this);
        },

        addResolvedRecord: function (record) {
            this.selectedItems.push(record);
            var internalRecord = new Office.Controls.PeoplePicker.internalPeoplePickerRecord(this, record);
            internalRecord.add();
            internalRecord.updateHoverText();
            this.internalSelectedItems.push(internalRecord);
            this.onDataSelected(record);
            this.onAdded(this, record.principalInfo);
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
            this.displayLoadingIcon(record.text);
            this.dataProvider.getPrincipals(record.DisplayName, function (error, ps) {
                if (ps !== null) {
                    internalRecord = self.onDataReceivedForResolve(ps, internalRecord);
                    self.onAdded(this, internalRecord.Record.principalInfo);
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
            this.displayLoadingIcon(input);
            this.dataProvider.getPrincipals(input, function (error, ps) {
                if (ps !== null) {
                    internalRecord = self.onDataReceivedForResolve(ps, internalRecord);
                    if (triggerUserListener) {
                        self.onAdded(self, internalRecord.Record.principalInfo);
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
            holderDiv.innerHTML = Office.Controls.peoplePickerTemplates.generateRecordTemplate(this.Record, this.parent.allowMultiple, this.parent.showImage);
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
            this.parent.defaultText.className = 'office-hide';
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
                this.principalOptions[options[i].PersonId] = options[i];
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
            if (key.keyCode === 38) {
                if (target.previousSibling !== null) {
                    this.parent.autofill.changeFocus(target, target.previousSibling);
                    target.previousSibling.focus();
                } else if (target.parentNode.parentNode.nextSibling !== null) {
                    var autofillElementsUlTags = this.parent.root.querySelectorAll('ul'),
                    ul = autofillElementsUlTags[1];
                    this.parent.autofill.changeFocus(target, ul.lastChild);
                    ul.lastChild.focus();
                } else {
                    var recentList = this.parent.root.querySelector('ul.ms-PeoplePicker-resultList');
                    this.parent.autofill.changeFocus(target, recentList.lastChild);
                    recentList.lastChild.focus();
                }
            } else if (key.keyCode === 40) {
                if (target.nextSibling !== null) {
                    this.parent.autofill.changeFocus(target, target.nextSibling);
                    target.nextSibling.focus();
                } else if (target.parentNode.parentNode.nextSibling !== null) {
                    var autofillElementsUlTags = this.parent.root.querySelectorAll('ul'),
                    ul = autofillElementsUlTags[1];
                    this.parent.autofill.changeFocus(target, ul.firstChild);
                    ul.firstChild.focus();
                }
            } else if (key.keyCode === 9) {
                var personId = this.parent.autofill.getPersonIdFromListElement(target);
                this.onAutofillClick(this.parent.autofill.entries[personId]);
                Office.Controls.Utils.cancelEvent(e);
            }
            return true;
        },

        resolveTo: function (principal) {
            Office.Controls.PeoplePicker.copyToRecord(this.Record, principal);
            this.Record.text = principal.DisplayName;
            this.Record.isResolved = true;
            if (this.parent.enableCache) {
                this.parent.addToCache(principal);
            }
            Office.Controls.Utils.removeClass(this.Node, 'has-error');
            var primaryTextNode = this.Node.querySelector('div.ms-Persona-primaryText');
            primaryTextNode.innerHTML = Office.Controls.Utils.htmlEncode(principal.DisplayName);
            this.updateHoverText();
        },

        refresh: function (principal) {
            Office.Controls.PeoplePicker.copyToRecord(this.Record, principal);
            this.Record.text = principal.DisplayName;
            var primaryTextNode = this.Node.querySelector('div.ms-Persona-primaryText');
            primaryTextNode.innerHTML = Office.Controls.Utils.htmlEncode(principal.DisplayName);
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
            this.parent.onRemoved(this.parent, this.Record.principalInfo);
            this.resolveTo(selectedPrincipal);
            this.parent.refreshInputField();
            this.principalOptions = null;
            this.optionsList = null;
            if (this.parent.enableCache) {
                this.parent.addToCache(selectedPrincipal);
            }
            this.parent.validateMultipleMatchError();
            this.parent.autofill.close();
            this.parent.onAdded(this.parent, this.Record.principalInfo);
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
                this.entries[entries[i].PersonId] = entries[i];
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
                    if (Office.Controls.Utils.isNullOrUndefined(this.entries[currentEntry.PersonId])) {
                        this.entries[entries[i].PersonId] = entries[i];
                        newServerEntries.push(currentEntry);
                    }
                }
                this.serverEntries = newServerEntries;
            } else {
                this.entries = {};
                var length = entries.length, i;
                for (i = 0; i < length; i++) {
                    this.entries[entries[i].PersonId] = entries[i];
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
                    var image = li.querySelector('img');
                    var personId = this.getPersonIdFromListElement(li);
                    (function (self, image, personId) {
                        self.parent.dataProvider.getImageAsync(personId, function (error, imgSrc) {
                            if (imgSrc != null) {
                                image.style.backgroundImage = "url('" + imgSrc + "')";
                                self.entries[personId].imgSrc = imgSrc;
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

        openSearchingLoadingStatus: function (searchingName) {
            this.root.innerHTML = Office.Controls.peoplePickerTemplates.generateSerachingLoadingTemplate();
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
            PersonId = this.getPersonIdFromListElement(listItem);
            handler(this.entries[PersonId]);
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
            if (key.keyCode === 38) {
                if (target.previousSibling !== null) {
                    this.changeFocus(target, target.previousSibling);
                } else if (target.parentNode.parentNode.nextSibling !== null) {
                    var autofillElementsUlTags = this.root.querySelectorAll('ul'),
                    ul = autofillElementsUlTags[1];
                    this.changeFocus(target, ul.lastChild);
                } else {
                    var recentList = this.root.querySelector('ul.ms-PeoplePicker-resultList');
                    this.changeFocus(target, recentList.lastChild);
                }
            } else if (key.keyCode === 40) {
                if (target.nextSibling !== null) {
                    this.changeFocus(target, target.nextSibling);
                } else if (target.parentNode.parentNode.nextSibling !== null) {
                    var autofillElementsUlTags = this.root.querySelectorAll('ul'),
                    ul = autofillElementsUlTags[1];
                    this.changeFocus(target, ul.firstChild);
                }
            }
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
                if (cacheEntry.PersonId === entry.PersonId) {
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
            var userNameKey = candidate.PersonId;
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
            if (Office.Controls.Utils.isNullOrUndefined(candidate.DisplayName)) {
                candidate.DisplayName = '';
            }
            if ((!userNameKey.toLowerCase().indexOf(key)) || (!emailKey.toLowerCase().indexOf(key)) || (!candidate.DisplayName.toLowerCase().indexOf(key))) {
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
        if (Office.Controls.PeoplePicker.res.hasOwnProperty("newName")) {
            return Office.Controls.PeoplePicker.res[newName];
        }
        return Office.Controls.Utils.getStringFromResource('PeoplePicker', stringName);
    };

    Office.Controls.peoplePickerTemplates.getDefaultText = function (allowMultiple) {
        if (allowMultiple) {
            return Office.Controls.peoplePickerTemplates.getString('PP_DefaultMessagePlural');
        }
        return Office.Controls.peoplePickerTemplates.getString('PP_DefaultMessage');
    };

    Office.Controls.peoplePickerTemplates.generateControlTemplate = function (inputName, allowMultiple, inputHint) {
        var defaultText;
        if (Office.Controls.Utils.isNullOrEmptyString(inputHint)) {
            defaultText = Office.Controls.Utils.htmlEncode(Office.Controls.peoplePickerTemplates.getDefaultText(allowMultiple));
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
        itemHtml = '<li tabindex=\"0\" class=\"ms-PeoplePicker-result\" data-office-peoplepicker-value=\"' + Office.Controls.Utils.htmlEncode(principal.PersonId) + '\" title=\"' + titleText + '\">';
        itemHtml += '<div class=\"ms-Persona ms-PersonaAdded\">';
        if (showImage) {
            if (isCached && !Office.Controls.Utils.isNullOrUndefined(principal.imgSrc)) {
                itemHtml += '<img class=\"ms-Persona-image\" style=\"background-image:url(\'' + principal.imgSrc + '\')\">';
            } else {
                itemHtml += '<img class=\"ms-Persona-image\">';
            }
        }
        itemHtml += '<div class=\"ms-Persona-details\">';
        itemHtml += '<a onclick=\"return false;\" href=\"#\" tabindex=\"-1\">';
        itemHtml += '<div class=\"ms-Persona-primaryText\" >' + Office.Controls.Utils.htmlEncode(principal.DisplayName) + '</div>';
        if (!Office.Controls.Utils.isNullOrEmptyString(principal.Description)) {
            itemHtml += '<div class=\"ms-Persona-secondaryText\" >' + Office.Controls.Utils.htmlEncode(principal.Description) + '</div>';
        }
        itemHtml += '</a></div></div></li>';
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

    Office.Controls.peoplePickerTemplates.generateRecordTemplate = function (record, allowMultiple, showImage) {
        var recordHtml;
        var userRecordClass = 'ms-PeoplePicker-persona';
        if (!allowMultiple) {
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
                recordHtml += '<img class=\"ms-Persona-image\">';
            } else  {
                recordHtml += '<img class=\"ms-Persona-image\" style=\"background-image:url(\'' + record.imgSrc + '\')\">';
            }
        }
        recordHtml += '<div class=\"ms-Persona-details\">';
        recordHtml += '<div class=\"ms-Persona-primaryText ms-Persona-primaryTextForResolvedUserAdded\">' + Office.Controls.Utils.htmlEncode(record.text);
        recordHtml += '</div></div></div>';
        if (showImage) {
            recordHtml += '<div class=\"ms-PeoplePicker-personaRemove ms-PeoplePicker-personaRemoveWithImage\">';
        } else {
            recordHtml += '<div class=\"ms-PeoplePicker-personaRemove ms-PeoplePicker-personaRemoveNoImage\">';
        }
        recordHtml += '<i tabindex=\"0\" class=\"ms-Icon ms-Icon--x ms-Icon-added\">';
        recordHtml += '</i></div>';
        recordHtml += '</div>';
        return recordHtml;
    };

    Office.Controls.peoplePickerTemplates.generateAlertNode = function () {
        var alertHtml = '<div role=\"alert\" class=\"office-peoplepicker-alert\">';
        alertHtml += '</div>';
        return alertHtml;
    };

    Office.Controls.Context = function (parameterObject) {
        if (typeof parameterObject !== 'object') {
            Office.Controls.Utils.errorConsole('Invalid parameters type');
            return;
        }
        var sharepointHost = parameterObject.HostUrl;
        if (Office.Controls.Utils.isNullOrUndefined(sharepointHost)) {
            var param = Office.Controls.Utils.getQueryStringParameter('SPHostUrl');
            if (!Office.Controls.Utils.isNullOrEmptyString(param)) {
                param = decodeURIComponent(param);
            }
            this.HostUrl = param;
        } else {
            this.HostUrl = sharepointHost;
        }
        this.HostUrl = this.HostUrl.toLocaleLowerCase();
    };
    Office.Controls.Context.prototype = {
        HostUrl: null
    };

    Office.Controls.Runtime = function () { };
    Office.Controls.Runtime.initialize = function (parameterObject) {
        Office.Controls.Runtime.context = new Office.Controls.Context(parameterObject);
    };

    Office.Controls.Utils = function () { };
    Office.Controls.Utils.deserializeJSON = function (data) {
        if (Office.Controls.Utils.isNullOrEmptyString(data)) {
            return {};
        }
        return JSON.parse(data);
    };
    Office.Controls.Utils.serializeJSON = function (obj) {
        return JSON.stringify(obj);
    };
    Office.Controls.Utils.isNullOrEmptyString = function (str) {
        var strNull = null;
        return str === strNull || str === undefined || !str.length;
    };
    Office.Controls.Utils.isNullOrUndefined = function (obj) {
        var objNull = null;
        return obj === objNull || obj === undefined;
    };
    Office.Controls.Utils.getQueryStringParameter = function (paramToRetrieve) {
        if (document.URL.split('?').length < 2) {
            return null;
        }
        var queryParameters = document.URL.split('?')[1].split('#')[0].split('&'), i;
        for (i = 0; i < queryParameters.length; i = i + 1) {
            var singleParam = queryParameters[i].split('=');
            if (singleParam[0].toLowerCase() === paramToRetrieve.toLowerCase()) {
                return singleParam[1];
            }
        }
        return null;
    };

    Office.Controls.Utils.logConsole = function (message) {
        console.log(message);
    };

    Office.Controls.Utils.warnConsole = function (message) {
        console.warn(message);
    };

    Office.Controls.Utils.errorConsole = function (message) {
        console.error(message);
    };

    Office.Controls.Utils.getObjectFromFullyQualifiedName = function (objectName) {
        var currentObject = window.self;
        var controlNameParts = objectName.split('.'), i;
        for (i = 0; i < controlNameParts.length; i++) {
            currentObject = currentObject[controlNameParts[i]];
            if (Office.Controls.Utils.isNullOrUndefined(currentObject)) {
                return null;
            }
        }
        return currentObject;
    };

    Office.Controls.Utils.getStringFromResource = function (controlName, stringName) {
        var resourceObjectName = 'Office.Controls.' + controlName + 'Resources', res,
        nonPreserveCase = stringName.charAt(0).toString().toLowerCase() + stringName.substr(1);
        resourceObjectName += 'Defaults';
        res = Office.Controls.Utils.getObjectFromFullyQualifiedName(resourceObjectName);
        if (!Office.Controls.Utils.isNullOrUndefined(res)) {
            return res[stringName];
        }
        return stringName;
    };

    Office.Controls.Utils.addEventListener = function (element, eventName, handler) {
        var h = function (e) {
            try {
                return handler(e);
            } catch (ex) {
                throw ex;
            }
        };
        if (!Office.Controls.Utils.isNullOrUndefined(element.addEventListener)) {
            element.addEventListener(eventName, h, false);
        } else if (!Office.Controls.Utils.isNullOrUndefined(element.attachEvent)) {
            element.attachEvent('on' + eventName, h);
        }
    };

    Office.Controls.Utils.getEvent = function (e) {
        return (Office.Controls.Utils.isNullOrUndefined(e)) ? window.event : e;
    };

    Office.Controls.Utils.getTarget = function (e) {
        return (Office.Controls.Utils.isNullOrUndefined(e.target)) ? e.srcElement : e.target;
    };

    Office.Controls.Utils.cancelEvent = function (e) {
        if (!Office.Controls.Utils.isNullOrUndefined(e.cancelBubble)) {
            e.cancelBubble = true;
        }
        if (!Office.Controls.Utils.isNullOrUndefined(e.stopPropagation)) {
            e.stopPropagation();
        }
        if (!Office.Controls.Utils.isNullOrUndefined(e.preventDefault)) {
            e.preventDefault();
        }
        if (!Office.Controls.Utils.isNullOrUndefined(e.returnValue)) {
            e.returnValue = false;
        }
        if (!Office.Controls.Utils.isNullOrUndefined(e.cancel)) {
            e.cancel = true;
        }
    };

    Office.Controls.Utils.addClass = function (elem, className) {
        if (elem.className !== '') {
            elem.className += ' ';
        }
        elem.className += className;
    };

    Office.Controls.Utils.removeClass = function (elem, className) {
        var regex = new RegExp('( |^)' + className + '( |$)');
        elem.className = elem.className.replace(regex, ' ').trim();
    };

    Office.Controls.Utils.containClass = function (elem, className) {
        return elem.className.indexOf(className) !== -1;
    };

    Office.Controls.Utils.cloneData = function (obj) {
        return Office.Controls.Utils.deserializeJSON(Office.Controls.Utils.serializeJSON(obj));
    };

    Office.Controls.Utils.formatString = function (format) {
        var args = [], $ai_8;
        for ($ai_8 = 1; $ai_8 < arguments.length; ++$ai_8) {
            args[$ai_8 - 1] = arguments[$ai_8];
        }
        var result = '';
        var i = 0;
        while (i < format.length) {
            var open = Office.Controls.Utils.findPlaceHolder(format, i, '{');
            if (open < 0) {
                result = result + format.substr(i);
                break;
            }
            var close = Office.Controls.Utils.findPlaceHolder(format, open, '}');
            if (close > open) {
                result = result + format.substr(i, open - i);
                var position = format.substr(open + 1, close - open - 1);
                var pos = parseInt(position);
                result = result + args[pos];
                i = close + 1;
            } else {
                Office.Controls.Utils.errorConsole('Invalid Operation');
                return null;
            }
        }
        return result;
    };

    Office.Controls.Utils.findPlaceHolder = function (format, start, ch) {
        var index = format.indexOf(ch, start);
        while (index >= 0 && index < format.length - 1 && format.charAt(index + 1) === ch) {
            start = index + 2;
            index = format.indexOf(ch, start);
        }
        return index;
    };

    Office.Controls.Utils.htmlEncode = function (value) {
        value = value.replace(new RegExp('&', 'g'), '&amp;');
        value = value.replace(new RegExp('\"', 'g'), '&quot;');
        value = value.replace(new RegExp('\'', 'g'), '&#39;');
        value = value.replace(new RegExp('<', 'g'), '&lt;');
        value = value.replace(new RegExp('>', 'g'), '&gt;');
        return value;
    };

    Office.Controls.Utils.getLocalizedCountValue = function (locText, intervals, count) {
        var ret = '';
        var locIndex = -1;
        var intervalsArray = intervals.split('||'), i, length;
        for (i = 0, length = intervalsArray.length; i < length; i++) {
            var interval = intervalsArray[i];
            if (Office.Controls.Utils.isNullOrEmptyString(interval)) {
                continue;
            }
            var subIntervalsArray = interval.split(','), k, subLength;
            for (k = 0, subLength = subIntervalsArray.length; k < subLength; k++) {
                var subInterval = subIntervalsArray[k];
                if (Office.Controls.Utils.isNullOrEmptyString(subInterval)) {
                    continue;
                }
                if (isNaN(Number(subInterval))) {
                    var range = subInterval.split('-');
                    if (Office.Controls.Utils.isNullOrUndefined(range) || range.length !== 2) {
                        continue;
                    }
                    var min;
                    var max;
                    if (range[0] === '') {
                        min = 0;
                    } else {
                        if (isNaN(Number(range[0]))) {
                            continue;
                        }
                        min = parseInt(range[0]);
                    }
                    if (count >= min) {
                        if (range[1] === '') {
                            locIndex = i;
                            break;
                        } else {
                            if (isNaN(Number(range[1]))) {
                                continue;
                            }
                            max = parseInt(range[1]);
                        }
                        if (count <= max) {
                            locIndex = i;
                            break;
                        }
                    }
                } else {
                    var exactNumber = parseInt(subInterval);
                    if (count === exactNumber) {
                        locIndex = i;
                        break;
                    }
                }
            }
            if (locIndex !== -1) {
                break;
            }
        }
        var locValues = locText.split('||');
        if (locIndex !== -1) {
            ret = locValues[locIndex];
        }
        return ret;
    };
    Office.Controls.Utils.NOP = function () { };

    Office.Controls.PeoplePickerAadDataProvider = function (authContext) {
        this.authContext = authContext;
    }

    Office.Controls.PeoplePickerAadDataProvider.prototype = {
        maxResult: 50,
        authContext: null,
        getPrincipals: function (input, callback) {

            var self = this;
            self.authContext.acquireToken(Office.Controls.Utils.aadGraphResourceId, function (error, token) {

                // Handle ADAL Errors
                if (error || !token) {
                    callback('Error', null);
                    return;
                }
                var parsed = self.authContext._extractIdToken(token);
                var tenant = '';
                if (parsed) {
                    if (parsed.hasOwnProperty('tid')) {
                        tenant = parsed.tid;
                    }
                }

                var xhr = new XMLHttpRequest();
                xhr.open('GET', 'https://graph.windows.net/' + tenant + '/users?api-version=1.5' + "&$filter=startswith(displayName," +
                    encodeURIComponent("'" + input + "') or ") + "startswith(userPrincipalName," + encodeURIComponent("'" + input + "')"), true);
                xhr.setRequestHeader('Content-Type', 'application/json');
                xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                xhr.onabort = xhr.onerror = xhr.ontimeout = function () {
                    callback('Error', null);
                };
                xhr.onload = function () {
                    if (xhr.status === 401) {
                        callback('Unauthorized', null);
                        return;
                    }
                    if (xhr.status !== 200) {
                        callback('Unknown error', null);
                        return;
                    }
                    var result = JSON.parse(xhr.responseText), people = [];
                    if (result["odata.error"] !== undefined) {
                        callback(result["odata.error"], null);
                        return;
                    }
                    result.value.forEach(
                        function (e) {
                            var person = {};
                            person.DisplayName = e.displayName;
                            person.Description = e.department;
                            person.PersonId = e.objectId;
                            people.push(person);
                        });
                    if (people.length > self.maxResult) {
                        people = people.slice(0, self.maxResult);
                    }
                    callback(null, people);
                };
                xhr.send('');
            });
        },

        getImageAsync: function (personId, callback) {

            var self = this;
            self.authContext.acquireToken(Office.Controls.Utils.aadGraphResourceId, function (error, token) {

                // Handle ADAL Errors
                if (error || !token) {
                    callback('Error', null);
                    return;
                }
                var parsed = self.authContext._extractIdToken(token);
                var tenant = '';
                if (parsed) {
                    if (parsed.hasOwnProperty('tid')) {
                        tenant = parsed.tid;
                    }
                }

                var xhr = new XMLHttpRequest();
                xhr.open('GET', 'https://graph.windows.net/' + tenant + '/users/' + personId + '/thumbnailPhoto?api-version=1.5');
                xhr.setRequestHeader('Content-Type', 'application/json');
                xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                xhr.responseType = "blob";
                xhr.onabort = xhr.onerror = xhr.ontimeout = function () {
                    callback('Error', null);
                };
                xhr.onload = function () {
                    if (xhr.status === 401) {
                        callback('Unauthorized', null);
                        return;
                    }
                    if (xhr.status !== 200) {
                        callback('Unknown error', null);
                        return;
                    }
                    var reader = new FileReader();
                    reader.addEventListener("loadend", function() {
                        callback(null, reader.result);
                    });
                    reader.readAsDataURL(xhr.response);
                };
                xhr.send('');
            });
        }
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
    if (Office.Controls.Context.registerClass) { Office.Controls.Context.registerClass('Office.Controls.Context'); }
    if (Office.Controls.Runtime.registerClass) { Office.Controls.Runtime.registerClass('Office.Controls.Runtime'); }
    if (Office.Controls.Utils.registerClass) { Office.Controls.Utils.registerClass('Office.Controls.Utils'); }
    Office.Controls.PeoplePicker.res = {};
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
    Office.Controls.Runtime.context = null;
    Office.Controls.Utils.oDataJSONAcceptString = 'application/json;odata=verbose';
    Office.Controls.Utils.clientTagHeaderName = 'X-ClientService-ClientTag';
    Office.Controls.Utils.aadGraphResourceId = '00000002-0000-0000-c000-000000000000';
})();

