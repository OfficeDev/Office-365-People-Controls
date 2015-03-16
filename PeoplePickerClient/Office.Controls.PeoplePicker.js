/*! Version=16.00.0549.000 */
(function () {

    if (window.Type && window.Type.registerNamespace) {
        Type.registerNamespace('Office.Controls');
    } else {
        if (typeof (window['Office']) == 'undefined') {
            window['Office'] = new Object(); window['Office'].__namespace = true;
        }
        if (typeof (window['Office']['Controls']) == 'undefined') {
            window['Office']['Controls'] = new Object(); window['Office']['Controls'].__namespace = true;
        }

    }


    Office.Controls.PrincipalInfo = function () { }


    Office.Controls.PeoplePickerRecord = function () {
    }
    Office.Controls.PeoplePickerRecord.prototype = {
        isResolved: false,
        text: null,
        displayName: null,
        Description: null,
        PersonId: null,
        principalInfo: null,
    }


    Office.Controls.PeoplePicker = function (root, dataProvider, parameterObject) {
        try {
            if (typeof (root) !== 'object' || typeof (dataProvider) !== 'object' || !Office.Controls.Utils.isNullOrUndefined(parameterObject) && typeof (parameterObject) !== 'object') {
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
                if (!Office.Controls.Utils.isNullOrEmptyString(parameterObject.inputHint)) {
                    this.inputHint = parameterObject.inputHint;
                }
                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.showValidationErrors)) {
                    this.showValidationErrors = (String(parameterObject.showValidationErrors) === "true");
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

                if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.res)) {
                    Office.Controls.PeoplePicker.res = parameterObject.res;
                }
            }

            this.currentTimerId = -1;
            this.selectedItems = new Array(0);
            this.internalSelectedItems = new Array(0);
            this.errors = new Array(0);
            if (this.enableCache == true) {
                Office.Controls.Runtime.initialize({ HostUrl: window.location.host });
                this.cache = Office.Controls.PeoplePicker.mruCache.getInstance();
            }

            this.renderControl();
            this.autofill = new Office.Controls.PeoplePicker.autofillContainer(this);
        }
        catch (ex) {
            throw ex;
        }
    }
    Office.Controls.PeoplePicker.copyToRecord = function (record, info) {
        record.DisplayName = info.DisplayName;
        record.Description = info.Description;
        record.PersonId = info.PersonId;
        record.principalInfo = info;
    }

    Office.Controls.PeoplePicker.parseUserPaste = function (content) {
        var openBracket = content.indexOf('<');
        var emailSep = content.indexOf('@', openBracket);
        var closeBracket = content.indexOf('>', emailSep);
        if (openBracket !== -1 && emailSep !== -1 && closeBracket !== -1) {
            return content.substring(openBracket + 1, closeBracket);
        }
        return content;
    }
    Office.Controls.PeoplePicker.getSearchBoxClass = function () {
        return 'ms-PeoplePicker-searchBox';
    }
    Office.Controls.PeoplePicker.nopAddRemove = function (p1, p2) {
    }
    Office.Controls.PeoplePicker.nopOperation = function (p1) {
    }
    Office.Controls.PeoplePicker.create = function (root, parameterObject) {
        return new Office.Controls.PeoplePicker(root, parameterObject);
    }
    Office.Controls.PeoplePicker.prototype = {
        allowMultiple: false,
        startSearchCharLength: 1,
        delaySearchInterval: 300,
        enableCache: true,
        inputHint: null,
        onAdded: Office.Controls.PeoplePicker.nopAddRemove,
        onRemoved: Office.Controls.PeoplePicker.nopAddRemove,
        onChange: Office.Controls.PeoplePicker.nopOperation,
        onFocus: Office.Controls.PeoplePicker.nopOperation,
        onBlur: Office.Controls.PeoplePicker.nopOperation,
        onError: null,
        dataProvider: null,
        showValidationErrors: true,
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
        alertDiv: null,
        lastSearchQuery: '',
        currentToken: null,
        widthSet: false,
        currentPrincipalsChoices: null,
        hasErrors: false,
        errorDisplayed: null,
        hasMultipleEntryValidationError: false,
        hasMultipleMatchValidationError: false,
        hasNoMatchValidationError: false,
        autofill: null,

        reset: function () {
            while (this.internalSelectedItems.length) {
                var record = this.internalSelectedItems[0];
                record.removeAndNotTriggerUserListener();
            }
            this.setTextInputDisplayStyle();
            this.validateMultipleMatchError();
            this.validateMultipleEntryAllowed();
            this.validateNoMatchError();
            this.clearInputField();
            this.clearCacheData();
            if (Office.Controls.PeoplePicker.autofillContainer.currentOpened) {
                Office.Controls.PeoplePicker.autofillContainer.currentOpened.close();
            }
            Office.Controls.PeoplePicker.autofillContainer.currentOpened = null;
            Office.Controls.PeoplePicker.autofillContainer.boolBodyHandlerAdded = false;
            this.autofill = new Office.Controls.PeoplePicker.autofillContainer(this)
            this.toggleDefaultText();

        },

        remove: function (entryToRemove) {
            var record = this.internalSelectedItems;
            for (var i = 0; i < record.length; i++) {
                if (record[i].Record.principalInfo === entryToRemove) {
                    var recordToRemove = record[i].Record;
                    record[i].removeAndNotTriggerUserListener();
                    this.onRemoved(this, recordToRemove.principalInfo);
                    this.validateMultipleMatchError();
                    this.validateMultipleEntryAllowed();
                    this.validateNoMatchError();
                    this.setTextInputDisplayStyle();
                    this.textInput.focus();
                    break;
                }
            }
        },

        add: function (p1, resolved) {
            if (typeof (p1) === 'string') {
                this.addThroughString(p1);
            }
            else {
                var record = new Office.Controls.PeoplePickerRecord();
                Office.Controls.PeoplePicker.copyToRecord(record, p1)
                record.text = p1.DisplayName;
                if (Office.Controls.Utils.isNullOrUndefined(resolved)) {
                    record.isResolved = false;
                    this.addThroughRecord(record, false);
                }
                else {
                    record.isResolved = resolved;
                    this.addThroughRecord(record, resolved);
                }
            }
        },

        getAddedPeople: function () {
            var record = this.internalSelectedItems;
            var addedPeople = [];
            for (var i = 0; i < record.length; i++) {
                addedPeople[i] = record[i].Record.principalInfo;
            }
            return addedPeople;
        },

        clearCacheData: function () {
            if (this.cache != null) {
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
            var $$t_6 = this, $$t_7 = this;
            this.dataProvider.getPrincipals(userEmail, function (error, principalsReceived) {
                if (principalsReceived != null) {
                    Office.Controls.PeoplePicker.copyToRecord(record, principalsReceived[0]);
                    userInfoHandler(record);
                }
                else {
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
            }
            else {
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
            var $$t_7 = this;
            Office.Controls.Utils.addEventListener(this.actualRoot, 'click', function (e) {
                return $$t_7.onPickerClick(e);
            });
            this.inputData = this.actualRoot.querySelector('input[type=\"hidden\"]');
            this.textInput = this.actualRoot.querySelector('input[type=\"text\"]');
            this.defaultText = this.actualRoot.querySelector('span.office-peoplepicker-default');
            this.resolvedListRoot = this.actualRoot.querySelector('div.office-peoplepicker-recordList');
            this.autofillElement = this.actualRoot.querySelector('.ms-PeoplePicker-results');
            this.alertDiv = this.actualRoot.querySelector('.office-peoplepicker-alert');
            var $$t_8 = this;
            Office.Controls.Utils.addEventListener(this.textInput, 'focus', function (e) {
                return $$t_8.onInputFocus(e);
            });
            var $$t_9 = this;
            Office.Controls.Utils.addEventListener(this.textInput, 'blur', function (e) {
                return $$t_9.onInputBlur(e);
            });
            var $$t_A = this;
            Office.Controls.Utils.addEventListener(this.textInput, 'keydown', function (e) {
                return $$t_A.onInputKeyDown(e);
            });
            var $$t_B = this;
            Office.Controls.Utils.addEventListener(this.textInput, 'input', function (e) {
                return $$t_B.onInput(e);
            });
            var $$t_C = this;
            Office.Controls.Utils.addEventListener(window.self, 'resize', function (e) {
                return $$t_C.onResize(e);
            });
            this.toggleDefaultText();
            if (!Office.Controls.Utils.isNullOrUndefined(this.inputTabindex)) {
                this.textInput.setAttribute('tabindex', this.inputTabindex);
            }
        },

        toggleDefaultText: function () {
            if (this.actualRoot.className.indexOf('office-peoplepicker-autofill-focus') === -1 && this.showInputHint && !this.selectedItems.length && !this.textInput.value.length) {
                this.defaultText.className = 'office-peoplepicker-default office-helper';
            }
            else {
                this.defaultText.className = 'office-hide';
            }
        },

        onResize: function (e) {
            this.toggleDefaultText();
            return true;
        },

        onInputKeyDown: function (e) {
            var keyEvent = Office.Controls.Utils.getEvent(e);
            if (keyEvent.keyCode === 27) {
                this.autofill.close();
            }
            else if (keyEvent.keyCode === 9 && this.autofill.IsDisplayed) {
                var focusElement = this.autofillElement.querySelector("li.ms-PeoplePicker-resultAddedForSelect");
                if (focusElement != null) {
                    var personId = this.autofill.getPersonIdFromListElement(focusElement);
                    this.addResolvedPrincipal(this.autofill.entries[personId]);
                    this.autofill.flushContent();
                    Office.Controls.Utils.cancelEvent(e);
                    return false;
                }
                else {
                    this.autofill.close();
                }
            }
            else if ((keyEvent.keyCode === 40 || keyEvent.keyCode === 38 ) && this.autofill.IsDisplayed) {
                this.autofill.onKeyDownFromInput(keyEvent);
                keyEvent.preventDefault();
                keyEvent.stopPropagation();
                Office.Controls.Utils.cancelEvent(e);
                return false;
            }
            else if (keyEvent.keyCode === 37 && this.internalSelectedItems.length) {
                this.resolvedListRoot.lastChild.focus();
                Office.Controls.Utils.cancelEvent(e);
                return false;
            }
            else if (keyEvent.keyCode === 8) {
                var shouldRemove = false;
                if (!Office.Controls.Utils.isNullOrUndefined(document.selection)) {
                    var range = document.selection.createRange();
                    var selectedText = range.text;
                    range.moveStart('character', -this.textInput.value.length);
                    var caretPos = range.text.length;
                    if (!selectedText.length && !caretPos) {
                        shouldRemove = true;
                    }
                }
                else {
                    var selectionStart = this.textInput.selectionStart;
                    var selectionEnd = this.textInput.selectionEnd;
                    if (!selectionStart && selectionStart === selectionEnd) {
                        shouldRemove = true;
                    }
                }
                if (shouldRemove && this.internalSelectedItems.length) {
                    this.internalSelectedItems[this.internalSelectedItems.length - 1].remove();
                    Office.Controls.Utils.cancelEvent(e);
                }
            }
            else if ((keyEvent.keyCode === 75 && keyEvent.ctrlKey) || (keyEvent.keyCode === 186) || (keyEvent.keyCode === 59) || (keyEvent.keyCode === 13)) {
                keyEvent.preventDefault();
                keyEvent.stopPropagation();
                this.cancelLastRequest();
                this.attemptResolveInput();
                Office.Controls.Utils.cancelEvent(e);
                return false;
            }
            else if ((keyEvent.keyCode === 86 && keyEvent.ctrlKey) || (keyEvent.keyCode === 186)) {
                this.cancelLastRequest();
                var $$t_C = this;
                window.setTimeout(function () {
                    $$t_C.textInput.value = Office.Controls.PeoplePicker.parseUserPaste($$t_C.textInput.value);
                    $$t_C.attemptResolveInput();
                }, 0);
                return true;
            }
            else if (keyEvent.keyCode === 13 && keyEvent.shiftKey) {
                var $$t_D = this;
                this.autofill.open(function (selectedPrincipal) {
                    $$t_D.addResolvedPrincipal(selectedPrincipal);
                });
            }
            else {
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
            var cachedEntries = this.cache.get(this.textInput.value, 5);
            this.autofill.setCachedEntries(cachedEntries);
            if (!cachedEntries.length) {
                return;
            }
            var $$t_2 = this;
            this.autofill.open(function (selectedPrincipal) {
                $$t_2.addResolvedPrincipal(selectedPrincipal);
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
            var currentValue = this.textInput.value;
            var $$t_7 = this;
            this.currentTimerId = window.setTimeout(function () {
                if (currentValue !== $$t_7.lastSearchQuery || $$t_7.startSearchCharLength == 0) {
                    $$t_7.lastSearchQuery = currentValue;
                    if (currentValue.length >= $$t_7.startSearchCharLength) {
                        $$t_7.searchingTimes++;
                        $$t_7.displayLoadingIcon(currentValue);
                        $$t_7.removeValidationError('ServerProblem');
                        var token = new Office.Controls.PeoplePicker.cancelToken();
                        $$t_7.currentToken = token;
                        $$t_7.dataProvider.getPrincipals($$t_7.textInput.value, function (error, principalsReceived) {
                            if (principalsReceived != null) {
                                if (!token.IsCanceled) {
                                    $$t_7.onDataReceived(principalsReceived);
                                }
                                else {
                                    $$t_7.hideLoadingIcon();
                                }
                            }
                            else {
                                $$t_7.onDataFetchError(error);
                            }

                        });
                    }
                    else {
                        $$t_7.autofill.close();
                    }
                    if ($$t_7.enableCache) {
                        $$t_7.displayCachedEntries();
                    }
                }
            }, $$t_7.delaySearchInterval);
        },

        onDataFetchError: function (message) {
            this.hideLoadingIcon();
            this.addValidationError(Office.Controls.PeoplePicker.ValidationError.createServerProblemError());
        },

        onDataReceived: function (principalsReceived) {
            this.currentPrincipalsChoices = {};
            for (var i = 0; i < principalsReceived.length; i++) {
                var principal = principalsReceived[i];
                this.currentPrincipalsChoices[principal.PersonId] = principal;
            }
            this.autofill.setServerEntries(principalsReceived);
            this.hideLoadingIcon();
            var $$t_4 = this;
            this.autofill.open(function (selectedPrincipal) {
                $$t_4.addResolvedPrincipal(selectedPrincipal);
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
            }
            else {
                this.textInput.focus();
                this.textInput.setSelectionRange(endPos, endPos);
            }
        },

        onInputFocus: function (e) {
            if (Office.Controls.Utils.isNullOrEmptyString(this.actualRoot.className)) {
                this.actualRoot.className = 'office-peoplepicker-autofill-focus';
            }
            else {
                this.actualRoot.className += ' office-peoplepicker-autofill-focus';
            }
            if (!this.widthSet) {
                this.setInputMaxWidth();
            }
            this.toggleDefaultText();
            this.onFocus(this);

            var $$_9 = this;
            if (this.startSearchCharLength == 0 && (this.allowMultiple == true || this.internalSelectedItems.length == 0)) {
                this.startQueryAfterDelay();
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
            this.validateMultipleEntryAllowed();
            this.clearInputField();
            this.refreshInputField();
        },

        onDataRemoved: function (selectedPrincipal) {
            this.refreshInputField();
            this.validateMultipleMatchError();
            this.validateMultipleEntryAllowed();
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
               // this.actualRoot.className = 'ms-PeoplePicker';
                this.textInput.className = 'ms-PeoplePicker-searchFieldAddedForSingleSelectionHidden';
                this.textInput.setAttribute('readonly', 'readonly');
            }
            else {
                this.textInput.removeAttribute('readonly');
                this.textInput.className = 'ms-PeoplePicker-searchField ms-PeoplePicker-searchFieldAdded';
            }
        },

        changeAlertMessage: function (message) {
            this.alertDiv.innerHTML = Office.Controls.Utils.htmlEncode(message);
        },

        displayLoadingIcon: function (searchingName) {
            this.changeAlertMessage(Office.Controls.peoplePickerTemplates.getString('PP_Searching'));
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
            }
            else {
                internalRecordToResolve.setResolveOptions(principalsReceived);
            }
            this.refreshInputField();
            return internalRecordToResolve;
        },

        onDataReceivedForStalenessCheck: function (principalsReceived, internalRecordToCheck) {
            if (principalsReceived.length === 1) {
                internalRecordToCheck.resolveTo(principalsReceived[0]);
            }
            else {
                internalRecordToCheck.unresolve();
                internalRecordToCheck.setResolveOptions(principalsReceived);
            }
            this.refreshInputField();
        },

        addResolvedPrincipal: function (principal) {
            var record = new Office.Controls.PeoplePickerRecord();
            Office.Controls.PeoplePicker.copyToRecord(record, principal);
            record.text = principal.DisplayName;
            record.isResolved = true;
            this.selectedItems.push(record);
            var internalRecord = new Office.Controls.PeoplePicker.internalPeoplePickerRecord(this, record);
            internalRecord.add();
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
            this.internalSelectedItems.push(internalRecord);
            this.onDataSelected(record);
            this.onAdded(this, record.principalInfo);
            this.currentPrincipalsChoices = null;
        },

        addUncertainPrincipal: function (record) {
            this.selectedItems.push(record);
            var internalRecord = new Office.Controls.PeoplePicker.internalPeoplePickerRecord(this, record);
            internalRecord.add();
            this.internalSelectedItems.push(internalRecord);
            this.setTextInputDisplayStyle();
            this.displayLoadingIcon(record.text);
            var $$t_5 = this, $$t_6 = this;
            this.dataProvider.getPrincipals(record.DisplayName, function (error, ps) {
                if (ps != null) {
                    internalRecord = $$t_5.onDataReceivedForResolve(ps, internalRecord);
                    $$t_5.onAdded(this, internalRecord.Record.principalInfo);
                    $$t_5.onChange($$t_5);
                }
                else {
                    $$t_6.onDataFetchError(error);
                }
            });
            this.validateMultipleEntryAllowed();
        },

        addUnresolvedPrincipal: function (input, triggerUserListener) {
            var record = new Office.Controls.PeoplePickerRecord();
            var principalInfo = new Office.Controls.PrincipalInfo();
            principalInfo.displayName = input;
            record.text = input;
            record.principalInfo = principalInfo;
            record.isResolved = false;
            var internalRecord = new Office.Controls.PeoplePicker.internalPeoplePickerRecord(this, record);
            internalRecord.add();
            this.selectedItems.push(record);
            this.internalSelectedItems.push(internalRecord);
            this.clearInputField();
            this.setTextInputDisplayStyle();
            this.displayLoadingIcon(input);
            var $$t_7 = this, $$t_8 = this;
            this.dataProvider.getPrincipals(input, function (error, ps) {
                if (ps != null) {
                    internalRecord = $$t_7.onDataReceivedForResolve(ps, internalRecord);
                    if (triggerUserListener) {
                        $$t_7.onAdded($$t_7, internalRecord.Record.principalInfo);
                        $$t_7.onChange($$t_7);
                    }
                }
                else {
                    $$t_8.onDataFetchError(error);
                }
            });
            this.validateMultipleEntryAllowed();
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
            for (var i = 0; i < this.errors.length; i++) {
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
            }
            else {
                this.displayValidationErrors();
            }
        },

        validateMultipleEntryAllowed: function () {
            if (!this.allowMultiple) {
                if (this.selectedItems.length > 1) {
                    if (!this.hasMultipleEntryValidationError) {
                        this.addValidationError(Office.Controls.PeoplePicker.ValidationError.createMultipleEntryError());
                        this.hasMultipleEntryValidationError = true;
                    }
                }
                else if (this.hasMultipleEntryValidationError) {
                    this.removeValidationError('MultipleEntry');
                    this.hasMultipleEntryValidationError = false;
                }
            }
        },

        validateMultipleMatchError: function () {
            var oldStatus = this.hasMultipleMatchValidationError;
            var newStatus = false;
            for (var i = 0; i < this.internalSelectedItems.length; i++) {
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
            var oldStatus = this.hasNoMatchValidationError;
            var newStatus = false;
            for (var i = 0; i < this.internalSelectedItems.length; i++) {
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
            }
            else {
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
    }


    Office.Controls.PeoplePicker.internalPeoplePickerRecord = function (parent, record) {
        this.parent = parent;
        this.Record = record;
    }
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
            var recordRemovalEvent = Office.Controls.Utils.getEvent(e);
            var target = Office.Controls.Utils.getTarget(recordRemovalEvent);
            this.remove();
            Office.Controls.Utils.cancelEvent(e);
            this.parent.autofill.close();
            return false;
        },

        onRecordRemovalKeyDown: function (e) {
            var recordRemovalEvent = Office.Controls.Utils.getEvent(e);
            var target = Office.Controls.Utils.getTarget(recordRemovalEvent);
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
            }
            else if (keyEvent.keyCode === 37) {
                if (this.Node.previousSibling != null) {
                    this.Node.previousSibling.focus();
                }
                Office.Controls.Utils.cancelEvent(e);
            }
            else if (keyEvent.keyCode === 39) {
                if (this.Node.nextSibling != null) {
                    this.Node.nextSibling.focus();
                }
                else {
                    this.parent.textInput.focus();
                }
                Office.Controls.Utils.cancelEvent(e);
            }
            return false;
        },

        add: function () {
            var holderDiv = document.createElement('div');
            holderDiv.innerHTML = Office.Controls.peoplePickerTemplates.generateRecordTemplate(this.Record, this.parent.allowMultiple);
            var recordElement = holderDiv.firstChild;
            var $$t_4 = this;
            Office.Controls.Utils.addEventListener(recordElement, 'keydown', function (e) {
                return $$t_4.onRecordKeyDown(e);
            });

            var removeButtonElement = recordElement.querySelector('div.ms-PeoplePicker-personaRemove');
            var $$t_5 = this;
            Office.Controls.Utils.addEventListener(removeButtonElement, 'click', function (e) {
                return $$t_5.onRecordRemovalClick(e);
            });
            var $$t_6 = this;
            Office.Controls.Utils.addEventListener(removeButtonElement, 'keydown', function (e) {
                return $$t_6.onRecordRemovalKeyDown(e);
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
            for (var i = 0; i < this.parent.internalSelectedItems.length; i++) {
                if (this.parent.internalSelectedItems[i] === this) {
                    this.parent.internalSelectedItems.splice(i, 1);
                }
            }
            for (var i = 0; i < this.parent.selectedItems.length; i++) {
                if (this.parent.selectedItems[i] === this.Record) {
                    this.parent.selectedItems.splice(i, 1);
                }
            }
        },

        setResolveOptions: function (options) {
            this.optionsList = options;
            this.principalOptions = {};
            for (var i = 0; i < options.length; i++) {
                this.principalOptions[options[i].PersonId] = options[i];
            }
            var $$t_3 = this;
            Office.Controls.Utils.addEventListener(this.Node, 'click', function (e) {
                return $$t_3.onUnresolvedUserClick(e);
            });
            this.parent.validateMultipleMatchError();
            this.parent.validateNoMatchError();
        },

        onUnresolvedUserClick: function (e) {
            e = Office.Controls.Utils.getEvent(e);
            this.parent.autofill.flushContent();
            this.parent.autofill.setServerEntries(this.optionsList);
            var $$t_2 = this;
            this.parent.autofill.open(function (selectedPrincipal) {
                $$t_2.onAutofillClick(selectedPrincipal);
            });
            this.addKeyListenerForAutofill();
            this.parent.autofill.focusOnFirstElement();
            Office.Controls.Utils.cancelEvent(e);
            return false;
        },

        addKeyListenerForAutofill: function () {
            var autofillElementsLiTags = this.parent.autofill.root.querySelectorAll('li');
            for (var i = 0; i < autofillElementsLiTags.length; i++) {
                var li = autofillElementsLiTags[i];
                var $$c_5 = this;
                Office.Controls.Utils.addEventListener(li, 'keydown', function (e) {
                    return $$c_5.onAutofillKeyDown(e);
                });
            }
        },

        onAutofillKeyDown: function (e) {
            var key = Office.Controls.Utils.getEvent(e);
            var target = Office.Controls.Utils.getTarget(key);
            if (key.keyCode === 38) {
                if (target.previousSibling != null) {
                    this.parent.autofill.changeFocus(target, target.previousSibling);
                    target.previousSibling.focus();
                }
                else if (target.parentNode.parentNode.nextSibling != null) {
                    var autofillElementsUlTags = this.parent.root.querySelectorAll('ul');
                    var ul = autofillElementsUlTags[1];
                    this.parent.autofill.changeFocus(target, ul.lastChild);
                    ul.lastChild.focus();
                }
                else {
                    var recentList = this.parent.root.querySelector('ul.ms-PeoplePicker-resultList');
                    this.parent.autofill.changeFocus(target, recentList.lastChild);
                    recentList.lastChild.focus();
                }
            }
            else if (key.keyCode === 40) {
                if (target.nextSibling != null) {
                    this.parent.autofill.changeFocus(target, target.nextSibling);
                    target.nextSibling.focus();
                }
                else if (target.parentNode.parentNode.nextSibling != null) {
                    var autofillElementsUlTags = this.parent.root.querySelectorAll('ul');
                    var ul = autofillElementsUlTags[1];
                    this.parent.autofill.changeFocus(target, ul.firstChild);
                    ul.firstChild.focus();
                }
            }
            else if (key.keyCode === 9) {
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
            this.updateHoverText(primaryTextNode);
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
            this.updateHoverText(primaryTextNode);
        },

        updateHoverText: function (userLabel) {
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
    }


    Office.Controls.PeoplePicker.autofillContainer = function (parent) {
        this.entries = {};
        this.cachedEntries = new Array(0);
        this.serverEntries = new Array(0);
        this.parent = parent;
        this.root = parent.autofillElement;
        if (!Office.Controls.PeoplePicker.autofillContainer.boolBodyHandlerAdded) {
            var $$t_2 = this;
            Office.Controls.Utils.addEventListener(document.body, 'click', function (e) {
                return Office.Controls.PeoplePicker.autofillContainer.bodyOnClick(e);
            });
            Office.Controls.PeoplePicker.autofillContainer.boolBodyHandlerAdded = true;
        }
    }
    Office.Controls.PeoplePicker.autofillContainer.getControlRootFromSubElement = function (element) {
        while (element && element.nodeName.toLowerCase() !== 'body') {
            if (element.className.indexOf('office office-peoplepicker') !== -1) {
                return element;
            }
            element = element.parentNode;
        }
        return null;
    }
    Office.Controls.PeoplePicker.autofillContainer.bodyOnClick = function (e) {
        if (!Office.Controls.PeoplePicker.autofillContainer.currentOpened) {
            return true;
        }
        var click = Office.Controls.Utils.getEvent(e);
        var target = Office.Controls.Utils.getTarget(click);
        var controlRoot = Office.Controls.PeoplePicker.autofillContainer.getControlRootFromSubElement(target);
        if (!target || controlRoot !== Office.Controls.PeoplePicker.autofillContainer.currentOpened.parent.root) {
            Office.Controls.PeoplePicker.autofillContainer.currentOpened.close();
        }
        return true;
    }
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
            var length = entries.length;
            for (var i = 0; i < length; i++) {
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
            if (this.parent.enableCache == true) {
                var newServerEntries = new Array(0);
                var length = entries.length;
                for (var i = 0; i < length; i++) {
                    var currentEntry = entries[i];
                    if (Office.Controls.Utils.isNullOrUndefined(this.entries[currentEntry.PersonId])) {
                        this.entries[entries[i].PersonId] = entries[i];
                        newServerEntries.push(currentEntry);
                    }
                }
                this.serverEntries = newServerEntries;
            }
            else {
                this.entries = {};
                var length = entries.length;
                for (var i = 0; i < length; i++) {
                    this.entries[entries[i].PersonId] = entries[i];
                }
                this.serverEntries = entries;
            }
        },

        renderList: function (handler) {
            var isTabKey = false;
            this.root.innerHTML = Office.Controls.peoplePickerTemplates.generateAutofillListTemplate(this.cachedEntries, this.serverEntries, 30);
            var autofillElementsLinkTags = this.root.querySelectorAll('a');
            for (var i = 0; i < autofillElementsLinkTags.length; i++) {
                var link = autofillElementsLinkTags[i];
                var $$t_A = this;
                Office.Controls.Utils.addEventListener(link, 'click', function (e) {
                    return $$t_A.onEntryClick(e, handler);
                });
                var $$t_C = this;
                Office.Controls.Utils.addEventListener(link, 'focus', function (e) {
                    return $$t_C.onEntryFocus(e);
                });
                var $$t_D = this;
                Office.Controls.Utils.addEventListener(link, 'blur', function (e) {
                    return $$t_D.onEntryBlur(e, isTabKey);
                });
            }
            var autofillElementsLiTags = this.root.querySelectorAll('li');
            for (var i = 0; i < autofillElementsLiTags.length; i++) {
                var li = autofillElementsLiTags[i];
                var $$t_B = this;
               /* Office.Controls.Utils.addEventListener(li, 'keydown', function (e) {
                    var key = Office.Controls.Utils.getEvent(e);
                    isTabKey = (key.keyCode === 9);
                    if (key.keyCode === 32 || key.keyCode === 13) {
                        e.preventDefault();
                        e.stopPropagation();
                        return $$t_B.onEntryClick(e, handler);
                    }
                    return $$t_B.onKeyDown(e);
                });*/
            }
            if (autofillElementsLiTags.length > 0) {
                Office.Controls.Utils.addClass(autofillElementsLiTags[0], 'ms-PeoplePicker-resultAddedForSelect');
            }

        },

        flushContent: function () {
            var entry = this.root.querySelectorAll('div.ms-PeoplePicker-resultGroups');
            for (var i = 0; i < entry.length; i++) {
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
            if ((this.cachedEntries.length + this.serverEntries.length) > 0) {
                this.parent.changeAlertMessage(Office.Controls.peoplePickerTemplates.getString('PP_SuggestionsAvailable'));
            }
            else {
                this.parent.changeAlertMessage(Office.Controls.peoplePickerTemplates.getString('PP_NoSuggestionsAvailable'));
            }
        },

        close: function () {
            this.IsDisplayed = false;
            Office.Controls.Utils.removeClass(this.parent.actualRoot, 'is-active');
            //console.log("close autofill!");
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
            var click = Office.Controls.Utils.getEvent(e);
            var target = Office.Controls.Utils.getTarget(click);
            target = this.getParentListItem(target);
            var PersonId = this.getPersonIdFromListElement(target);
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
            var target =  this.root.querySelector("li.ms-PeoplePicker-resultAddedForSelect");
            if (key.keyCode === 38 ) {
                if (target.previousSibling != null) {
                    this.changeFocus(target, target.previousSibling);
                }
                else if (target.parentNode.parentNode.nextSibling != null) {
                    var autofillElementsUlTags = this.root.querySelectorAll('ul');
                    var ul = autofillElementsUlTags[1];
                    this.changeFocus(target, ul.lastChild);
                }
                else {
                    var recentList = this.root.querySelector('ul.ms-PeoplePicker-resultList');
                    this.changeFocus(target, recentList.lastChild);
                }
            }
            else if (key.keyCode === 40 ) {
                if (target.nextSibling != null) {
                    this.changeFocus(target, target.nextSibling);
                }
                else if (target.parentNode.parentNode.nextSibling != null) {
                    var autofillElementsUlTags = this.root.querySelectorAll('ul');
                    var ul = autofillElementsUlTags[1];
                    this.changeFocus(target, ul.firstChild);
                }
            }
            return true;
        },

        changeFocus : function (lastElement, nextElement){
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
                if ((next) && (next.nextSibling.className.toLowerCase() === 'ms-PeoplePicker-searchMore js-searchMore'.toLowerCase())) {
                    Office.Controls.PeoplePicker.autofillContainer.currentOpened.close();
                }
            }
            return false;
        }
    }

    Office.Controls.PeoplePicker.Parameters = function () { }

    Office.Controls.PeoplePicker.cancelToken = function () {
        this.IsCanceled = false;
    }
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
    }

    Office.Controls.PeoplePicker.ValidationError = function () {
    }
    Office.Controls.PeoplePicker.ValidationError.createMultipleMatchError = function () {
        var err = new Office.Controls.PeoplePicker.ValidationError();
        err.errorName = 'MultipleMatch';
        err.localizedErrorMessage = Office.Controls.peoplePickerTemplates.getString('PP_MultipleMatch');
        return err;
    }
    Office.Controls.PeoplePicker.ValidationError.createMultipleEntryError = function () {
        var err = new Office.Controls.PeoplePicker.ValidationError();
        err.errorName = 'MultipleEntry';
        err.localizedErrorMessage = Office.Controls.peoplePickerTemplates.getString('PP_MultipleEntry');
        return err;
    }
    Office.Controls.PeoplePicker.ValidationError.createNoMatchError = function () {
        var err = new Office.Controls.PeoplePicker.ValidationError();
        err.errorName = 'NoMatch';
        err.localizedErrorMessage = Office.Controls.peoplePickerTemplates.getString('PP_NoMatch');
        return err;
    }
    Office.Controls.PeoplePicker.ValidationError.createServerProblemError = function () {
        var err = new Office.Controls.PeoplePicker.ValidationError();
        err.errorName = 'ServerProblem';
        err.localizedErrorMessage = Office.Controls.peoplePickerTemplates.getString('PP_ServerProblem');
        return err;
    }
    Office.Controls.PeoplePicker.ValidationError.prototype = {
        errorName: null,
        localizedErrorMessage: null
    }


    Office.Controls.PeoplePicker.mruCache = function () {
        this.isCacheAvailable = this.checkCacheAvailability();
        if (!this.isCacheAvailable) {
            return;
        }

        this.initializeCache();
    }
    Office.Controls.PeoplePicker.mruCache.getInstance = function () {
        if (!Office.Controls.PeoplePicker.mruCache.instance) {
            Office.Controls.PeoplePicker.mruCache.instance = new Office.Controls.PeoplePicker.mruCache();
        }
        return Office.Controls.PeoplePicker.mruCache.instance;
    }
    Office.Controls.PeoplePicker.mruCache.prototype = {
        isCacheAvailable: false,
        _localStorage: null,
        _dataObject: null,

        get: function (key, maxResults) {
            if (Office.Controls.Utils.isNullOrUndefined(maxResults) || !maxResults) {
                maxResults = 2147483647;
            }
            var numberOfResults = 0;
            var results = new Array(0);
            var cache = this.dataObject.cacheMapping[Office.Controls.Runtime.context.HostUrl];
            var cacheLength = cache.length;
            for (var i = cacheLength; i > 0 && numberOfResults < maxResults; i--) {
                var candidate = cache[i - 1];
                if (this.entityMatches(candidate, key)) {
                    results.push(candidate);
                    numberOfResults += 1;
                }
            }
            return results;
        },

        set: function (entry) {
            var cache = this.dataObject.cacheMapping[Office.Controls.Runtime.context.HostUrl];
            var cacheSize = cache.length;
            var alreadyThere = false;
            for (var i = 0; i < cacheSize; i++) {
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
            if (!userNameKey.toLowerCase().indexOf(key) || !emailKey.toLowerCase().indexOf(key) || !candidate.DisplayName.toLowerCase().indexOf(key)) {
                return true;
            }
            return false;
        },

        initializeCache: function () {
            var cacheData = this.cacheRetreive('Office.PeoplePicker.Cache');
            if (Office.Controls.Utils.isNullOrEmptyString(cacheData)) {
                this.dataObject = new Office.Controls.PeoplePicker.mruCache.mruData();
            }
            else {
                var datas = Office.Controls.Utils.deserializeJSON(cacheData);
                if (datas.cacheVersion) {
                    this.dataObject = new Office.Controls.PeoplePicker.mruCache.mruData();
                    this.cacheDelete('Office.PeoplePicker.Cache');
                }
                else {
                    this.dataObject = datas;
                }
            }
            if (Office.Controls.Utils.isNullOrUndefined(this.dataObject.cacheMapping[Office.Controls.Runtime.context.HostUrl])) {
                this.dataObject.cacheMapping[Office.Controls.Runtime.context.HostUrl] = new Array(0);
            }
        },

        checkCacheAvailability: function () {
            try {
                if (typeof window.self.localStorage == 'undefined') {
                    return false;
                }
                else {
                    this.localStorage = window.self.localStorage;
                    return true;
                }
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
    }


    Office.Controls.PeoplePicker.mruCache.mruData = function () {
        this.cacheMapping = {};
        this.cacheVersion = 0;
    }

    Office.Controls.PeoplePickerResourcesDefaults = function () {
    }


    Office.Controls.peoplePickerTemplates = function () {
    }
    Office.Controls.peoplePickerTemplates.getString = function (stringName) {
        var newName = 'PeoplePicker' + stringName.substr(3);
        if ((newName) in Office.Controls.PeoplePicker.res) {
            return Office.Controls.PeoplePicker.res[newName];
        }
        else {
            return Office.Controls.Utils.getStringFromResource('PeoplePicker', stringName);
        }
    }
    Office.Controls.peoplePickerTemplates.getDefaultText = function (allowMultiple) {
        if (allowMultiple) {
            return Office.Controls.peoplePickerTemplates.getString('PP_DefaultMessagePlural');
        }
        else {
            return Office.Controls.peoplePickerTemplates.getString('PP_DefaultMessage');
        }
    }
    Office.Controls.peoplePickerTemplates.generateControlTemplate = function (inputName, allowMultiple, inputHint) {
        var defaultText;
        if (Office.Controls.Utils.isNullOrEmptyString(inputHint)) {
            defaultText = Office.Controls.Utils.htmlEncode(Office.Controls.peoplePickerTemplates.getDefaultText(allowMultiple));
        }
        else {
            defaultText = Office.Controls.Utils.htmlEncode(inputHint);
        }
        var body = '<div class=\"ms-PeoplePicker\">';
        body += '<input type=\"hidden\"';
        if (!Office.Controls.Utils.isNullOrEmptyString(inputName)) {
            body += ' name=\"' + Office.Controls.Utils.htmlEncode(inputName) + '\"';
        }
        body += '/>';
        body += '<div class=\"ms-PeoplePicker-searchBox ms-PeoplePicker-searchBoxAdded\">';
        body += '<span class=\"office-peoplepicker-default office-helper\">' + defaultText + '</span>';
        body += '<div class=\"office-peoplepicker-recordList\"></div>';
        body += '<input type=\"text\" class=\"ms-PeoplePicker-searchField ms-PeoplePicker-searchFieldAdded\" size=\"1\" autocorrect=\"off\" autocomplete=\"off\" autocapitalize=\"off\"  style=\"width:106px;\" />';
        body += '</div>';
        body += '<div class=\"ms-PeoplePicker-results\">';
        body += '</div>';
        body += Office.Controls.peoplePickerTemplates.generateAlertNode();
        body += '</div>';
        return body;
    }
    Office.Controls.peoplePickerTemplates.generateErrorTemplate = function (ErrorMessage) {
        var innerHtml = '<span class=\"office-peoplepicker-error office-error\">';
        innerHtml += Office.Controls.Utils.htmlEncode(ErrorMessage);
        innerHtml += '</span>';
        return innerHtml;
    }
    Office.Controls.peoplePickerTemplates.generateAutofillListItemTemplate = function (principal, source) {
        var titleText = Office.Controls.Utils.htmlEncode((Office.Controls.Utils.isNullOrEmptyString(principal.Email)) ? '' : principal.Email);
        var itemHtml = '<li tabindex=\"0\" class=\"ms-PeoplePicker-result\" data-office-peoplepicker-value=\"' + Office.Controls.Utils.htmlEncode(principal.PersonId) + '\" title=\"' + titleText + '\">';
        itemHtml += '<div  class=\"ms-Persona ms-PersonaAdded\">';
        itemHtml += '<div  class=\"ms-Persona-details ms-Persona-detailsForDropdownAdded\">';
        itemHtml += '<a onclick=\"return false;\" href=\"#\" tabindex=\"-1\">';
        itemHtml += '<div class=\"ms-Persona-primaryText ms-Persona-primaryTextAdded\" >' + Office.Controls.Utils.htmlEncode(principal.DisplayName) + '</div>';
        if (!Office.Controls.Utils.isNullOrEmptyString(principal.Description)) {
            itemHtml += '<div class=\"ms-Persona-secondaryText ms-Persona-secondaryTextAdded\" >' + Office.Controls.Utils.htmlEncode(principal.Description) + '</div>';
        }
        itemHtml += '</a></div></div></li>';
        return itemHtml;
    }
    Office.Controls.peoplePickerTemplates.generateAutofillListTemplate = function (cachedEntries, serverEntries, maxCount) {
        var html = '<div class=\"ms-PeoplePicker-resultGroups\">';
        if (Office.Controls.Utils.isNullOrUndefined(cachedEntries)) {
            cachedEntries = new Array(0);
        }
        if (Office.Controls.Utils.isNullOrUndefined(serverEntries)) {
            serverEntries = new Array(0);
        }
        html += Office.Controls.peoplePickerTemplates.generateAutofillGroupTemplate(cachedEntries, 1, true);
        html += Office.Controls.peoplePickerTemplates.generateAutofillGroupTemplate(serverEntries, 0, false);
        html += '</div>';
        html += Office.Controls.peoplePickerTemplates.generateAutofillFooterTemplate(cachedEntries.length + serverEntries.length, maxCount);
        return html;
    }
    Office.Controls.peoplePickerTemplates.generateAutofillGroupTemplate = function (principals, source, isCached) {
        var listHtml = '';
        if (!principals.length) {
            return listHtml;
        }
        var cachedGrouptTitile = Office.Controls.Utils.htmlEncode(Office.Controls.peoplePickerTemplates.getString('PP_SearchResultRecentGroup'));
        var searchedGroupTitile = Office.Controls.Utils.htmlEncode(Office.Controls.peoplePickerTemplates.getString('PP_SearchResultMoreGroup'));
        var groupTitle = (isCached) ? cachedGrouptTitile : searchedGroupTitile;
        listHtml += '<div class=\"ms-PeoplePicker-resultGroup\">';
        listHtml += '<div class=\"ms-PeoplePicker-resultGroupTitle ms-PeoplePicker-resultGroupTitleAdded\">' + groupTitle + '</div>';
        listHtml += '<ul class=\"ms-PeoplePicker-resultList\" id=\"' + groupTitle + '\">';
        for (var i = 0; i < principals.length; i++) {
            listHtml += Office.Controls.peoplePickerTemplates.generateAutofillListItemTemplate(principals[i], source);
        }
        listHtml += '</ul></div>';
        return listHtml;
    }
    Office.Controls.peoplePickerTemplates.generateAutofillFooterTemplate = function (count, maxCount) {
        var footerHtml = '<div class=\"ms-PeoplePicker-searchMore js-searchMore\">';
        footerHtml += '<div class=\"ms-PeoplePicker-searchMoreIcon\"></div>';
        var footerText;
        if (count >= maxCount) {
            footerText = Office.Controls.Utils.formatString(Office.Controls.peoplePickerTemplates.getString('PP_ShowingTopNumberOfResults'), count.toString());
        }
        else {
            footerText = Office.Controls.Utils.formatString(Office.Controls.Utils.getLocalizedCountValue(Office.Controls.peoplePickerTemplates.getString('PP_Results'), Office.Controls.peoplePickerTemplates.getString('PP_ResultsIntervals'), count), count.toString());
        }
        footerText = Office.Controls.Utils.htmlEncode(footerText);
        footerHtml += '<div class=\"ms-PeoplePicker-searchMorePrimary ms-PeoplePicker-searchMorePrimaryAdded\">' + footerText + '</div>';
        footerHtml += '</div>';
        return footerHtml;
    }
    Office.Controls.peoplePickerTemplates.generateSerachingLoadingTemplate = function () {
        var searchingLable = Office.Controls.Utils.htmlEncode(Office.Controls.peoplePickerTemplates.getString('PP_Searching'));
        var searchingLoadingHtml = '<div class=\"ms-PeoplePicker-searchMore js-searchMore is-searching\">';
        searchingLoadingHtml += '<div class=\"ms-PeoplePicker-searchMoreIconFixed\"></div>';
        searchingLoadingHtml += '<div class=\"ms-PeoplePicker-searchMorePrimary ms-PeoplePicker-searchMorePrimaryAdded\">' + searchingLable + '</div>';
        searchingLoadingHtml += '</div>';
        return searchingLoadingHtml;
    }
    Office.Controls.peoplePickerTemplates.generateRecordTemplate = function (record, allowMultiple) {
        var recordHtml;
        var userRecordClass = 'ms-PeoplePicker-persona';
        if (!allowMultiple) {
            userRecordClass += ' ms-PeoplePicker-personaForSingleAdded';
        }
        if (record.isResolved) {
            recordHtml = '<div class=\"' + userRecordClass + '\" tabindex=\"0\">';
        }
        else {
            recordHtml = '<div class=\"' + userRecordClass + ' ' + 'has-error' + '\" tabindex=\"0\">';
        }
        recordHtml += '<div class=\"ms-Persona ms-Persona--xs ms-PersonaAddedForRecord\" >';
        recordHtml += '<div class=\"ms-Persona-details ms-Persona-detailsAdded\">';
        recordHtml += '<div class=\"ms-Persona-primaryText ms-Persona-primaryTextForResolvedUserAdded\">' + Office.Controls.Utils.htmlEncode(record.text);
        recordHtml += '</div></div></div>';
        recordHtml += '<div class=\"ms-PeoplePicker-personaRemove ms-PeoplePicker-personaRemoveAdded\">';
        recordHtml += '<i tabindex=\"0\" class=\"ms-Icon ms-Icon--x ms-Icon-added\">';
        recordHtml += '</i></div>';
        recordHtml += '</div>';
        return recordHtml;
    }
    Office.Controls.peoplePickerTemplates.generateAlertNode = function () {
        var alertHtml = '<div role=\"alert\" class=\"office-peoplepicker-alert\">';
        alertHtml += '</div>';
        return alertHtml;
    }


    Office.Controls.Context = function (parameterObject) {
        if (typeof (parameterObject) !== 'object') {
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
        }
        else {
            this.HostUrl = sharepointHost;
        }
        this.HostUrl = this.HostUrl.toLocaleLowerCase();
    }
    Office.Controls.Context.prototype = {
        HostUrl: null,
    }


    Office.Controls.Runtime = function () {
    }
    Office.Controls.Runtime.initialize = function (parameterObject) {
        Office.Controls.Runtime.context = new Office.Controls.Context(parameterObject);
    }


    Office.Controls.Utils = function () {
    }
    Office.Controls.Utils.deserializeJSON = function (data) {
        if (Office.Controls.Utils.isNullOrEmptyString(data)) {
            return {};
        }
        else {
            return JSON.parse(data);
        }
    }
    Office.Controls.Utils.serializeJSON = function (obj) {
        return JSON.stringify(obj);
    }
    Office.Controls.Utils.isNullOrEmptyString = function (str) {
        var strNull = null;
        return str === strNull || typeof (str) === 'undefined' || !str.length;
    }
    Office.Controls.Utils.isNullOrUndefined = function (obj) {
        var objNull = null;
        return obj === objNull || typeof (obj) === 'undefined';
    }
    Office.Controls.Utils.getQueryStringParameter = function (paramToRetrieve) {
        if (document.URL.split('?').length < 2) {
            return null;
        }
        var queryParameters = document.URL.split('?')[1].split('#')[0].split('&');
        for (var i = 0; i < queryParameters.length; i = i + 1) {
            var singleParam = queryParameters[i].split('=');
            if (singleParam[0].toLowerCase() === paramToRetrieve.toLowerCase()) {
                return singleParam[1];
            }
        }
        return null;
    }
    Office.Controls.Utils.logConsole = function (message) {
        console.log(message);
    }
    Office.Controls.Utils.warnConsole = function (message) {
        console.warn(message);
    }
    Office.Controls.Utils.errorConsole = function (message) {
        console.error(message);
    }
    Office.Controls.Utils.getObjectFromFullyQualifiedName = function (objectName) {
        var currentObject = window.self;
        var controlNameParts = objectName.split('.');
        for (var i = 0; i < controlNameParts.length; i++) {
            currentObject = currentObject[controlNameParts[i]];
            if (Office.Controls.Utils.isNullOrUndefined(currentObject)) {
                return null;
            }
        }
        return currentObject;
    }
    Office.Controls.Utils.getStringFromResource = function (controlName, stringName) {
        var resourceObjectName = 'Office.Controls.' + controlName + 'Resources';
        var res;
        var nonPreserveCase = stringName.charAt(0).toString().toLowerCase() + stringName.substr(1);
        resourceObjectName += 'Defaults';
        res = Office.Controls.Utils.getObjectFromFullyQualifiedName(resourceObjectName);
        if (!Office.Controls.Utils.isNullOrUndefined(res)) {
            return res[stringName];
        }
        return stringName;
    }
    Office.Controls.Utils.addEventListener = function (element, eventName, handler) {
        var h = function (e) {
            try {
                return handler(e);
            }
            catch (ex) {
                throw ex;
            }
        };
        if (!Office.Controls.Utils.isNullOrUndefined(element.addEventListener)) {
            element.addEventListener(eventName, h, false);
        }
        else if (!Office.Controls.Utils.isNullOrUndefined(element.attachEvent)) {
            element.attachEvent('on' + eventName, h);
        }
    }
    Office.Controls.Utils.getEvent = function (e) {
        return (Office.Controls.Utils.isNullOrUndefined(e)) ? window.event : e;
    }
    Office.Controls.Utils.getTarget = function (e) {
        return (Office.Controls.Utils.isNullOrUndefined(e.target)) ? e.srcElement : e.target;
    }
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
    }
    Office.Controls.Utils.addClass = function (elem, className) {
        if (elem.className !== '') {
            elem.className += ' ';
        }
        elem.className += className;
    }
    Office.Controls.Utils.removeClass = function (elem, className) {
        var regex = new RegExp('( |^)' + className + '( |$)');
        elem.className = elem.className.replace(regex, ' ').trim();
    }
    Office.Controls.Utils.containClass = function (elem, className) {
        return elem.className.indexOf(className) !== -1;
    }
    Office.Controls.Utils.cloneData = function (obj) {
        return Office.Controls.Utils.deserializeJSON(Office.Controls.Utils.serializeJSON(obj));
    }
    Office.Controls.Utils.formatString = function (format) {
        var args = [];
        for (var $ai_8 = 1; $ai_8 < arguments.length; ++$ai_8) {
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
            else {
                var close = Office.Controls.Utils.findPlaceHolder(format, open, '}');
                if (close > open) {
                    result = result + format.substr(i, open - i);
                    var position = format.substr(open + 1, close - open - 1);
                    var pos = parseInt(position);
                    result = result + args[pos];
                    i = close + 1;
                }
                else {
                    Office.Controls.Utils.errorConsole('Invalid Operation');
                    return null;
                }
            }
        }
        return result;
    }
    Office.Controls.Utils.findPlaceHolder = function (format, start, ch) {
        var index = format.indexOf(ch, start);
        while (index >= 0 && index < format.length - 1 && format.charAt(index + 1) === ch) {
            start = index + 2;
            index = format.indexOf(ch, start);
        }
        return index;
    }
    Office.Controls.Utils.htmlEncode = function (value) {
        value = value.replace(new RegExp('&', 'g'), '&amp;');
        value = value.replace(new RegExp('\"', 'g'), '&quot;');
        value = value.replace(new RegExp('\'', 'g'), '&#39;');
        value = value.replace(new RegExp('<', 'g'), '&lt;');
        value = value.replace(new RegExp('>', 'g'), '&gt;');
        return value;
    }
    Office.Controls.Utils.getLocalizedCountValue = function (locText, intervals, count) {
        var ret = '';
        var locIndex = -1;
        var intervalsArray = intervals.split('||');
        for (var i = 0, lenght = intervalsArray.length; i < lenght; i++) {
            var interval = intervalsArray[i];
            if (Office.Controls.Utils.isNullOrEmptyString(interval)) {
                continue;
            }
            var subIntervalsArray = interval.split(',');
            for (var k = 0, subLenght = subIntervalsArray.length; k < subLenght; k++) {
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
                    }
                    else {
                        if (isNaN(Number(range[0]))) {
                            continue;
                        }
                        else {
                            min = parseInt(range[0]);
                        }
                    }
                    if (count >= min) {
                        if (range[1] === '') {
                            locIndex = i;
                            break;
                        }
                        else {
                            if (isNaN(Number(range[1]))) {
                                continue;
                            }
                            else {
                                max = parseInt(range[1]);
                            }
                        }
                        if (count <= max) {
                            locIndex = i;
                            break;
                        }
                    }
                }
                else {
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
    }
    Office.Controls.Utils.NOP = function () {
    }


    if (Office.Controls.PrincipalInfo.registerClass) Office.Controls.PrincipalInfo.registerClass('Office.Controls.PrincipalInfo');
    if (Office.Controls.PeoplePickerRecord.registerClass) Office.Controls.PeoplePickerRecord.registerClass('Office.Controls.PeoplePickerRecord');
    if (Office.Controls.PeoplePicker.registerClass) Office.Controls.PeoplePicker.registerClass('Office.Controls.PeoplePicker');
    if (Office.Controls.PeoplePicker.internalPeoplePickerRecord.registerClass) Office.Controls.PeoplePicker.internalPeoplePickerRecord.registerClass('Office.Controls.PeoplePicker.internalPeoplePickerRecord');
    if (Office.Controls.PeoplePicker.autofillContainer.registerClass) Office.Controls.PeoplePicker.autofillContainer.registerClass('Office.Controls.PeoplePicker.autofillContainer');
    if (Office.Controls.PeoplePicker.Parameters.registerClass) Office.Controls.PeoplePicker.Parameters.registerClass('Office.Controls.PeoplePicker.Parameters');
    if (Office.Controls.PeoplePicker.cancelToken.registerClass) Office.Controls.PeoplePicker.cancelToken.registerClass('Office.Controls.PeoplePicker.cancelToken');
    if (Office.Controls.PeoplePicker.ValidationError.registerClass) Office.Controls.PeoplePicker.ValidationError.registerClass('Office.Controls.PeoplePicker.ValidationError');
    if (Office.Controls.PeoplePicker.mruCache.registerClass) Office.Controls.PeoplePicker.mruCache.registerClass('Office.Controls.PeoplePicker.mruCache');
    if (Office.Controls.PeoplePicker.mruCache.mruData.registerClass) Office.Controls.PeoplePicker.mruCache.mruData.registerClass('Office.Controls.PeoplePicker.mruCache.mruData');
    if (Office.Controls.PeoplePickerResourcesDefaults.registerClass) Office.Controls.PeoplePickerResourcesDefaults.registerClass('Office.Controls.PeoplePickerResourcesDefaults');
    if (Office.Controls.peoplePickerTemplates.registerClass) Office.Controls.peoplePickerTemplates.registerClass('Office.Controls.peoplePickerTemplates');
    if (Office.Controls.Context.registerClass) Office.Controls.Context.registerClass('Office.Controls.Context');
    if (Office.Controls.Runtime.registerClass) Office.Controls.Runtime.registerClass('Office.Controls.Runtime');
    if (Office.Controls.Utils.registerClass) Office.Controls.Utils.registerClass('Office.Controls.Utils');
    Office.Controls.PeoplePicker.res = {};
    Office.Controls.PeoplePicker.autofillContainer.currentOpened = null;
    Office.Controls.PeoplePicker.autofillContainer.boolBodyHandlerAdded = false;
    Office.Controls.PeoplePicker.ValidationError.multipleMatchName = 'MultipleMatch';
    Office.Controls.PeoplePicker.ValidationError.multipleEntryName = 'MultipleEntry';
    Office.Controls.PeoplePicker.ValidationError.noMatchName = 'NoMatch';
    Office.Controls.PeoplePicker.ValidationError.serverProblemName = 'ServerProblem';
    Office.Controls.PeoplePicker.mruCache.instance = null;
    Office.Controls.PeoplePickerResourcesDefaults.PP_SuggestionsAvailable = 'Suggestions Available';
    Office.Controls.PeoplePickerResourcesDefaults.PP_NoMatch = 'We couldn\'t find an exact match.';
    Office.Controls.PeoplePickerResourcesDefaults.PP_ShowingTopNumberOfResults = '{0} found';
    Office.Controls.PeoplePickerResourcesDefaults.PP_ServerProblem = 'Sorry, we\'re having trouble reaching the server.';
    Office.Controls.PeoplePickerResourcesDefaults.PP_DefaultMessagePlural = 'Enter names or email addresses...';
    Office.Controls.PeoplePickerResourcesDefaults.PP_MultipleMatch = 'Multiple entries matched, please click to resolve.';
    Office.Controls.PeoplePickerResourcesDefaults.PP_Results = 'No results found||{0} found||{0} found';
    Office.Controls.PeoplePickerResourcesDefaults.PP_Searching = 'Searching...';
    Office.Controls.PeoplePickerResourcesDefaults.PP_ResultsIntervals = '0||1||2-';
    Office.Controls.PeoplePickerResourcesDefaults.PP_NoSuggestionsAvailable = 'No Suggestions Available';
    Office.Controls.PeoplePickerResourcesDefaults.PP_RemovePerson = 'Remove person or group {0}';
    Office.Controls.PeoplePickerResourcesDefaults.PP_DefaultMessage = 'Enter a name or email address...';
    Office.Controls.PeoplePickerResourcesDefaults.PP_MultipleEntry = 'You can only enter one name.';
    Office.Controls.PeoplePickerResourcesDefaults.PP_SearchResultRecentGroup = 'Recent';
    Office.Controls.PeoplePickerResourcesDefaults.PP_SearchResultMoreGroup = 'More';
    Office.Controls.Runtime.context = null;
    Office.Controls.Utils.oDataJSONAcceptString = 'application/json;odata=verbose';
    Office.Controls.Utils.clientTagHeaderName = 'X-ClientService-ClientTag';
})();
