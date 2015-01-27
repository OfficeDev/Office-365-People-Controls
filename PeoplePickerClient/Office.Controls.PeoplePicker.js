/*! Version=16.00.0549.000 */
(function(){

Type.registerNamespace('Office.Controls');

Office.Controls.PrincipalInfo = function() {}


Office.Controls.PeoplePickerRecord = function() {
}
Office.Controls.PeoplePickerRecord.prototype = {
    isResolved: false,
    text: null,
    displayName: null,
    Description: null,
    PersonId: null,
    principalInfo: null,
}


Office.Controls.PeoplePicker = function(root, dataProvider, parameterObject) {
    this._currentTimerId = -1;
    this.selectedItems = new Array(0);
    this._internalSelectedItems = new Array(0);
    this.errors = new Array(0);
    this._cache = Office.Controls.PeoplePicker._mruCache.getInstance();
    try {
        if (typeof(root) !== 'object' || typeof(dataProvider) !== 'object' ||typeof(parameterObject) !== 'object') {
            Office.Controls.Utils.errorConsole('Invalid parameters type');
            return;
        }
        this._root = root;
        this._dataProvider = dataProvider;

        if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.allowMultipleSelections)) {
            this._allowMultiple = parameterObject.allowMultipleSelections;
        }
        if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.startSearchCharLength)) {
            this.startSearchCharLength = parameterObject.startSearchCharLength;
        }
        if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.delaySearchInterval)) {
            this.delaySearchInterval = parameterObject.delaySearchInterval;
        }
        if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.enableCache)) {
            this.enableCache = parameterObject.enableCache;
        }
        if (!Office.Controls.Utils.isNullOrEmptyString(parameterObject.inputHint)) {
            this.inputHint = parameterObject.inputHint;
        }
        if (!Office.Controls.Utils.isNullOrUndefined(parameterObject.showValidationErrors)) {
            this._showValidationErrors = parameterObject.showValidationErrors;
        }

        this._onAdded = parameterObject.onAdded;
        if (Office.Controls.Utils.isNullOrUndefined(this._onAdded)) {
            this._onAdded = Office.Controls.PeoplePicker._nopAddRemove;
        }
        this._onRemoved = parameterObject.onRemoved;
        if (Office.Controls.Utils.isNullOrUndefined(this._onRemoved)) {
            this._onRemoved = Office.Controls.PeoplePicker._nopAddRemove;
        }
        this._onChange = parameterObject.onChange;
        if (Office.Controls.Utils.isNullOrUndefined(this._onChange)) {
            this._onChange = Office.Controls.PeoplePicker._nopOperation;
        }
        this._onFocus = parameterObject.onFocus;
        if (Office.Controls.Utils.isNullOrUndefined(this._onFocus)) {
            this._onFocus = Office.Controls.PeoplePicker._nopOperation;
        }
        this._onBlur = parameterObject.onBlur;
        if (Office.Controls.Utils.isNullOrUndefined(this._onBlur)) {
            this._onBlur = Office.Controls.PeoplePicker._nopOperation;
        }
        this._onError = parameterObject.onError;

        Office.Controls.PeoplePicker._res$i = parameterObject.res;
        if (Office.Controls.Utils.isNullOrUndefined(Office.Controls.PeoplePicker._res$i)) {
            Office.Controls.PeoplePicker._res$i = {};
        }

        this._renderControl();
        this._autofill = new Office.Controls.PeoplePicker._autofillContainer(this);
    }
    catch (ex) {
        throw ex;
    }
}
Office.Controls.PeoplePicker._copyToRecord$i = function(record, info) {
    record.displayName = info.DisplayName;
    record.Description = info.Description;
    record.PersonId = info.PersonId;
    record.principal = info;
}

Office.Controls.PeoplePicker._parseUserPaste = function(content) {
    var openBracket = content.indexOf('<');
    var emailSep = content.indexOf('@', openBracket);
    var closeBracket = content.indexOf('>', emailSep);
    if (openBracket !== -1 && emailSep !== -1 && closeBracket !== -1) {
        return content.substring(openBracket + 1, closeBracket);
    }
    return content;
}
Office.Controls.PeoplePicker.getSearchBoxClass = function() {
    return 'ms-PeoplePicker-searchBox';
}
Office.Controls.PeoplePicker._nopAddRemove = function(p1, p2) {
}
Office.Controls.PeoplePicker._nopOperation = function(p1) {
}
Office.Controls.PeoplePicker.create = function(root, parameterObject) {
    return new Office.Controls.PeoplePicker(root, parameterObject);
}
Office.Controls.PeoplePicker.prototype = {
    _allowMultiple: false,
    startSearchCharLength: 3,
    delaySearchInterval: 300,
    enableCache: true,
    inputHint: null,
    _onAdded: null,
    _onRemoved: null,
    _onChange: null,
    _onFocus: null,
    _onBlur: null,
    _onError: null,
    _dataProvider: null,
    _showValidationErrors: true,
    _showInputHint: true,
    _inputTabindex: 0,
    _searchingTimes: 0,
    _inputBeginAction: false,
    _actualRoot: null,
    _textInput: null,
    _inputData: null,
    _defaultText: null,
    _resolvedListRoot: null,
    _autofillElement: null,
    _errorMessageElement: null,
    _root: null,
    _alertDiv: null,
    _lastSearchQuery: '',
    _currentToken: null,
    _widthSet: false,
    _currentPrincipalsChoices: null,
    hasErrors: false,
    _errorDisplayed: null,
    _hasMultipleEntryValidationError: false,
    _hasMultipleMatchValidationError: false,
    _hasNoMatchValidationError: false,
    _autofill: null,
    
    reset: function() {
        while (this._internalSelectedItems.length) {
            var record = this._internalSelectedItems[0];
            record._removeAndNotTriggerUserListener();
        }
        this._setTextInputDisplayStyle();
        this._validateMultipleMatchError();
        this._validateMultipleEntryAllowed();
        this._validateNoMatchError();
        this._clearInputField();
        if (Office.Controls.PeoplePicker._autofillContainer.currentOpened) {
            Office.Controls.PeoplePicker._autofillContainer.currentOpened.close();
        }
        this._toggleDefaultText();
    },
    
    remove: function(entryToRemove) {
        var record = this._internalSelectedItems;
        for (var i = 0; i < record.length; i++) {
            if (record[i].Record === entryToRemove) {
                record[i]._removeAndNotTriggerUserListener();
                break;
            }
        }
    },
    
    add: function(p1, resolve) {
        if (typeof(p1) === 'string') {
            this._addThroughString(p1);
        }
        else {
            var record = new Office.Controls.PeoplePickerRecord();
            Office.Controls.PeoplePicker._copyToRecord$i(record, p1)
            if (Office.Controls.Utils.isNullOrUndefined(resolve)) {
                this._addThroughRecord(record, false);
            }
            else {
                this._addThroughRecord(record, resolve);
            }
        }
    },

    getAddedPeople: function () {
        var record = this._internalSelectedItems;
        var addedPeople = {}
        for (var i = 0; i < record.length; i++) {
            addedPeople[i]= record[i].record.info;
        }
        return addedPeople;
    },

    getErrorDisplayed: function () {
        return this._errorDisplayed;
    },
    
    getUserInfoAsync: function(userInfoHandler, userEmail) {
        var record = new Office.Controls.PeoplePickerRecord();
        var $$t_6 = this, $$t_7 = this;
        this._dataProvider.getPrincipals(userEmail, function(principalsReceived) {
            Office.Controls.PeoplePicker._copyToRecord$i(record, principalsReceived[0]);
            userInfoHandler(record);
        }, function(error) {
            userInfoHandler(null);
        });
    },
    
    get_textInput: function() {
        return this._textInput;
    },
    
    get_actualRoot: function() {
        return this._actualRoot;
    },
    
    _addThroughString: function(input) {
        if (Office.Controls.Utils.isNullOrEmptyString(input)) {
            Office.Controls.Utils.errorConsole('Input can\'t be null or empty string. PeoplePicker Id : ' + this._root.id);
            return;
        }
        this._addUnresolvedPrincipal(input, false);
    },
    
    _addThroughRecord: function(info, resolve) {
        if (resolve) {
            this._addUncertainPrincipal(info);
        }
        else {
            this._addResolvedRecord(info);
        }
    },
    
    _renderControl: function(inputName) {
        this._root.innerHTML = Office.Controls._peoplePickerTemplates.generateControlTemplate(inputName, this._allowMultiple, this.inputHint);
        if (this._root.className.length > 0) {
            this._root.className += ' ';
        }
        this._root.className += 'office office-peoplepicker';
        this._actualRoot = this._root.querySelector('div.ms-PeoplePicker');
        var $$t_7 = this;
        Office.Controls.Utils.addEventListener(this._actualRoot, 'click', function(e) {
            return $$t_7._onPickerClick(e);
        });
        this._inputData = this._actualRoot.querySelector('input[type=\"hidden\"]');
        this._textInput = this._actualRoot.querySelector('input[type=\"text\"]');
        var $$t_8 = this;
        Office.Controls.Utils.addEventListener(this._textInput, 'focus', function(e) {
            return $$t_8._onInputFocus(e);
        });
        var $$t_9 = this;
        Office.Controls.Utils.addEventListener(this._textInput, 'blur', function(e) {
            return $$t_9._onInputBlur(e);
        });
        var $$t_A = this;
        Office.Controls.Utils.addEventListener(this._textInput, 'keydown', function(e) {
            return $$t_A._onInputKeyDown(e);
        });
        var $$t_B = this;
        Office.Controls.Utils.addEventListener(this._textInput, 'keyup', function(e) {
            return $$t_B._onInputKeyUp(e);
        });
        var $$t_C = this;
        Office.Controls.Utils.addEventListener(window.self, 'resize', function(e) {
            return $$t_C._onResize(e);
        });
        this._defaultText = this._actualRoot.querySelector('span.office-peoplepicker-default');
        this._resolvedListRoot = this._actualRoot.querySelector('div.office-peoplepicker-recordList');
        this._autofillElement = this._actualRoot.querySelector('.ms-PeoplePicker-results');
        this._alertDiv = this._actualRoot.querySelector('.office-peoplepicker-alert');
        this._toggleDefaultText();
        if (!Office.Controls.Utils.isNullOrUndefined(this._inputTabindex)) {
            this._textInput.setAttribute('tabindex', this._inputTabindex);
        }
    },
    
    _toggleDefaultText: function() {
        if ( this._actualRoot.className.indexOf('office-peoplepicker-autofill-focus') === -1 && this._showInputHint && !this.selectedItems.length && !this._textInput.value.length) {
            this._defaultText.className = 'office-peoplepicker-default office-helper';
        }
        else {
            this._defaultText.className = 'office-hide';
        }
    },
    
    _onResize: function(e) {
        this._toggleDefaultText();
        return true;
    },
    
    _onInputKeyDown: function(e) {
        var keyEvent = Office.Controls.Utils.getEvent(e);
        if (keyEvent.keyCode === 27) {
            this._autofill.close();
        }
        else if (keyEvent.keyCode === 40 && this._autofill.IsDisplayed) {
            var firstElement = this._autofillElement.querySelector('a');
            if (firstElement && firstElement.firstChild) {
                firstElement.firstChild.focus();
                Office.Controls.Utils.cancelEvent(e);
            }
        }
        else if (keyEvent.keyCode === 8) {
            var shouldRemove = false;
            if (!Office.Controls.Utils.isNullOrUndefined(document.selection)) {
                var range = document.selection.createRange();
                var selectedText = range.text;
                range.moveStart('character', -this._textInput.value.length);
                var caretPos = range.text.length;
                if (!selectedText.length && !caretPos) {
                    shouldRemove = true;
                }
            }
            else {
                var selectionStart = this._textInput.selectionStart;
                var selectionEnd = this._textInput.selectionEnd;
                if (!selectionStart && selectionStart === selectionEnd) {
                    shouldRemove = true;
                }
            }
            if (shouldRemove && this._internalSelectedItems.length) {
                this._internalSelectedItems[this._internalSelectedItems.length - 1]._remove();
            }
        }
        else if ((keyEvent.keyCode === 75 && keyEvent.ctrlKey) || (keyEvent.keyCode === 186) || (keyEvent.keyCode === 9 && this._autofill.IsDisplayed) || (keyEvent.keyCode === 13)) {
            keyEvent.preventDefault();
            keyEvent.stopPropagation();
            this._cancelLastRequest();
            this._attemptResolveInput();
            Office.Controls.Utils.cancelEvent(e);
            return false;
        }
        else if ((keyEvent.keyCode === 86 && keyEvent.ctrlKey) || (keyEvent.keyCode === 186)) {
            this._cancelLastRequest();
            var $$t_C = this;
            window.setTimeout(function() {
                $$t_C._textInput.value = Office.Controls.PeoplePicker._parseUserPaste($$t_C._textInput.value);
                $$t_C._attemptResolveInput();
            }, 0);
            return true;
        }
        else if (keyEvent.keyCode === 13 && keyEvent.shiftKey) {
            var $$t_D = this;
            this._autofill.open(function(selectedPrincipal) {
                $$t_D._addResolvedPrincipal(selectedPrincipal);
            });
        }
        else {
            this._resizeInputField();
        }
        return true;
    },
    
    _cancelLastRequest: function() {
        window.clearTimeout(this._currentTimerId);
        if (!Office.Controls.Utils.isNullOrUndefined(this._currentToken)) {
            this._hideLoadingIcon();
            this._currentToken.cancel();
            this._currentToken = null;
        }
    },
    
    _onInputKeyUp: function(e) {
        this._startQueryAfterDelay();
        this._resizeInputField();
        this._autofill.close();
        return true;
    },
    
    _displayCachedEntries: function() {
        var cachedEntries = this._cache.get(this._textInput.value, 5);
        this._autofill.setCachedEntries(cachedEntries);
        if (!cachedEntries.length) {
            return;
        }
        var $$t_2 = this;
        this._autofill.open(function(selectedPrincipal) {
            $$t_2._addResolvedPrincipal(selectedPrincipal);
        });
    },
    
    _resizeInputField: function() {
        var size = Math.max(this._textInput.value.length + 1, 1);
        this._textInput.size = size;
    },
    
    _clearInputField: function() {
        this._textInput.value = '';
        this._resizeInputField();
    },
    
    _startQueryAfterDelay: function() {
        this._cancelLastRequest();
        var currentValue = this._textInput.value;
        var $$t_7 = this;
        this._currentTimerId = window.setTimeout(function() {
            if (currentValue !== $$t_7._lastSearchQuery) {
                $$t_7._lastSearchQuery = currentValue;
                if (currentValue.length >= $$t_7.startSearchCharLength) {
                    $$t_7._searchingTimes++;
                    $$t_7._displayLoadingIcon(currentValue);
                    $$t_7._removeValidationError('ServerProblem');
                    var token = new Office.Controls.PeoplePicker._cancelToken();
                    $$t_7._currentToken = token;
                    $$t_7._dataProvider.getPrincipals($$t_7._textInput.value, function(principalsReceived) {
                        if (!token.IsCanceled) {
                            $$t_7._onDataReceived(principalsReceived);
                        }
                        else {
                            $$t_7._hideLoadingIcon();
                        }
                    }, function(error) {
                        $$t_7._onDataFetchError(error);
                    });
                }
                else {
                    $$t_7._autofill.close();
                }
                if ($$t_7.enableCache) {
                    $$t_7._displayCachedEntries();
                }
            }
        }, $$t_7.delaySearchInterval);
    },
    
    _onDataFetchError: function(message) {
        this._hideLoadingIcon();
        this._addValidationError(Office.Controls.PeoplePicker.ValidationError._createServerProblemError$i());
    },
    
    _onDataReceived: function(principalsReceived) {
        this._currentPrincipalsChoices = {};
        for (var i = 0; i < principalsReceived.length; i++) {
            var principal = principalsReceived[i];
            this._currentPrincipalsChoices[principal.PersonId] = principal;
        }
        this._autofill.setServerEntries(principalsReceived);
        this._hideLoadingIcon();
        var $$t_4 = this;
        this._autofill.open(function(selectedPrincipal) {
            $$t_4._addResolvedPrincipal(selectedPrincipal);
        });
    },
    
    _onPickerClick: function(e) {
        this._textInput.focus();
        e = Office.Controls.Utils.getEvent(e);
        var element = Office.Controls.Utils.getTarget(e);
        if (element.nodeName.toLowerCase() !== 'input') {
            this._focusToEnd();
        }
        return true;
    },
    
    _focusToEnd: function() {
        var endPos = this._textInput.value.length;
        if (!Office.Controls.Utils.isNullOrUndefined(this._textInput.createTextRange)) {
            var range = this._textInput.createTextRange();
            range.collapse(true);
            range.moveStart('character', endPos);
            range.moveEnd('character', endPos);
            range.select();
        }
        else {
            this._textInput.focus();
            this._textInput.setSelectionRange(endPos, endPos);
        }
    },
    
    _onInputFocus: function(e) {
        if (Office.Controls.Utils.isNullOrEmptyString(this._actualRoot.className)) {
            this._actualRoot.className = 'office-peoplepicker-autofill-focus';
        }
        else {
            this._actualRoot.className += ' office-peoplepicker-autofill-focus';
        }
        if (!this._widthSet) {
            this._setInputMaxWidth();
        }
        this._toggleDefaultText();
        this._onFocus(this);
        return true;
    },
    
    _setInputMaxWidth: function() {
        var maxwidth = this._actualRoot.clientWidth - 25;
        if (maxwidth <= 0) {
            maxwidth = 20;
        }
        this._textInput.style.maxWidth = maxwidth.toString() + 'px';
        this._widthSet = true;
    },
    
    _onInputBlur: function(e) {
        Office.Controls.Utils.removeClass(this._actualRoot, 'office-peoplepicker-autofill-focus');
        if (this._textInput.value.length > 0 || this.selectedItems.length > 0) {
            this._onBlur(this);
            return true;
        }
        this._toggleDefaultText();
        this._onBlur(this);
        return true;
    },
    
    _onDataSelected: function(selectedPrincipal) {
        this._lastSearchQuery = '';
        this._validateMultipleEntryAllowed();
        this._clearInputField();
        this._refreshInputField();
    },
    
    _onDataRemoved: function(selectedPrincipal) {
        this._refreshInputField();
        this._validateMultipleMatchError();
        this._validateMultipleEntryAllowed();
        this._validateNoMatchError();
        this._onRemoved(this, selectedPrincipal.info);
        this._onChange(this);
    },
    
    _addToCache: function(entry) {
        if (!this._cache.isCacheAvailable) {
            return;
        }
        this._cache.set(entry);
    },
    
    _refreshInputField: function() {
        this._inputData.value = Office.Controls.Utils.serializeJSON(this.selectedItems);
        this._setTextInputDisplayStyle();
    },
    
    _setTextInputDisplayStyle: function() {
        if ((!this._allowMultiple) && (this._internalSelectedItems.length === 1)) {
            this._actualRoot.className = 'ms-PeoplePicker';
            this._textInput.className = 'ms-PeoplePicker-searchFieldAddedForSingleSelectionHidden';
            this._textInput.setAttribute('readonly', 'readonly');
        }
        else {
            this._textInput.removeAttribute('readonly');
            this._textInput.className = 'ms-PeoplePicker-searchField ms-PeoplePicker-searchFieldAdded';
        }
    },
    
    _changeAlertMessage: function(message) {
        this._alertDiv.innerHTML = Office.Controls.Utils.htmlEncode(message);
    },
    
    _displayLoadingIcon: function(searchingName) {
        this._changeAlertMessage(Office.Controls._peoplePickerTemplates.getString('PP_Searching'));
        this._autofill.openSearchingLoadingStatus(searchingName);
    },
    
    _hideLoadingIcon: function() {
        this._autofill.closeSearchingLoadingStatus();
    },
    
    _attemptResolveInput: function() {
        this._autofill.close();
        if (this._textInput.value.length > 0) {
            this._lastSearchQuery = '';
            this._addUnresolvedPrincipal(this._textInput.value, true);
            this._clearInputField();
        }
    },
    
    _onDataReceivedForResolve: function(principalsReceived, internalRecordToResolve) {
        this._hideLoadingIcon();
        if (principalsReceived.length === 1) {
            internalRecordToResolve._resolveTo(principalsReceived[0]);
        }
        else {
            internalRecordToResolve._setResolveOptions(principalsReceived);
        }
        this._refreshInputField();
        return internalRecordToResolve;
    },
    
    _onDataReceivedForStalenessCheck: function(principalsReceived, internalRecordToCheck) {
        if (principalsReceived.length === 1) {
            internalRecordToCheck._refresh(principalsReceived[0]);
        }
        else {
            internalRecordToCheck._unresolve();
            internalRecordToCheck._setResolveOptions(principalsReceived);
        }
        this._refreshInputField();
    },
    
    _addResolvedPrincipal: function(principal) {
        var record = new Office.Controls.PeoplePickerRecord();
        Office.Controls.PeoplePicker._copyToRecord$i(record, principal);
        record.text = principal.DisplayName;
        record.isResolved = true;
        this.selectedItems.push(record);
        var internalRecord = new Office.Controls.PeoplePicker._internalPeoplePickerRecord(this, record);
        internalRecord._add();
        this._internalSelectedItems.push(internalRecord);
        this._onDataSelected(record);
        if (this.enableCache) {
            this._addToCache(principal);
        }
        this._currentPrincipalsChoices = null;
        this._autofill.close();
        this._textInput.focus();
        this._onAdded(this, record.info);
        this._onChange(this);
    },
    
    _addResolvedRecord: function(record) {
        this.selectedItems.push(record);
        var internalRecord = new Office.Controls.PeoplePicker._internalPeoplePickerRecord(this, record);
        internalRecord._add();
        this._internalSelectedItems.push(internalRecord);
        this._onDataSelected(record);
        this._onAdded(this, record.info);
        this._currentPrincipalsChoices = null;
    },
    
    _addUncertainPrincipal: function(record) {
        this.selectedItems.push(record);
        var internalRecord = new Office.Controls.PeoplePicker._internalPeoplePickerRecord(this, record);
        internalRecord._add();
        this._internalSelectedItems.push(internalRecord);
        var $$t_5 = this, $$t_6 = this;
        this._dataProvider.getPrincipals(record.email, function(ps) {
            $$t_5._onDataReceivedForStalenessCheck(ps, internalRecord);
            this._onAdded(this, record.info);
        }, function(message) {
            $$t_6._onDataFetchError(message);
        });
        this._validateMultipleEntryAllowed();
    },
    
    _addUnresolvedPrincipal: function(input, triggerUserListener) {
        var record = new Office.Controls.PeoplePickerRecord();
        var principalInfo = new Office.controls.PrincipalInfo();
        principalInfo.displayName = input;
        record.text = input;
        record.info = principalInfo;
        record.isResolved = false;
        var internalRecord = new Office.Controls.PeoplePicker._internalPeoplePickerRecord(this, record);
        internalRecord._add();
        this.selectedItems.push(record);
        this._internalSelectedItems.push(internalRecord);
        this._displayLoadingIcon(input);
        var $$t_7 = this, $$t_8 = this;
        this._dataProvider.getPrincipals(input, function(ps) {
            internalRecord = $$t_7._onDataReceivedForResolve(ps, internalRecord);
            if (triggerUserListener) {
                $$t_7._onAdded($$t_7, internalRecord.Record.info);
                $$t_7._onChange($$t_7);
            }
        }, function(message) {
            $$t_8._onDataFetchError(message);
        });
        this._validateMultipleEntryAllowed();
    },
    
    _addValidationError: function(err) {
        this.hasErrors = true;
        this.errors.push(err);
        this._displayValidationErrors();
        if (!Office.Controls.Utils.isNullOrUndefined(this._onError)) {
            this._onError(this, error);
        }
    },
    
    _removeValidationError: function(errorName) {
        for (var i = 0; i < this.errors.length; i++) {
            if (this.errors[i].errorName === errorName) {
                this.errors.splice(i, 1);
                break;
            }
        }
        if (!this.errors.length) {
            this.hasErrors = false;
        }
        if (!Office.Controls.Utils.isNullOrUndefined(this._onError) && this.errors.length) {
            this._onError(this, this.errors[0]);
        }
        else {
            this._displayValidationErrors();
        }
    },
    
    _validateMultipleEntryAllowed: function() {
        if (!this._allowMultiple) {
            if (this.selectedItems.length > 1) {
                if (!this._hasMultipleEntryValidationError) {
                    this._addValidationError(Office.Controls.PeoplePicker.ValidationError._createMultipleEntryError$i());
                    this._hasMultipleEntryValidationError = true;
                }
            }
            else if (this._hasMultipleEntryValidationError) {
                this._removeValidationError('MultipleEntry');
                this._hasMultipleEntryValidationError = false;
            }
        }
    },
    
    _validateMultipleMatchError: function() {
        var oldStatus = this._hasMultipleMatchValidationError;
        var newStatus = false;
        for (var i = 0; i < this._internalSelectedItems.length; i++) {
            if (!Office.Controls.Utils.isNullOrUndefined(this._internalSelectedItems[i]._optionsList) && this._internalSelectedItems[i]._optionsList.length > 0) {
                newStatus = true;
                break;
            }
        }
        if (!oldStatus && newStatus) {
            this._addValidationError(Office.Controls.PeoplePicker.ValidationError._createMultipleMatchError$i());
        }
        if (oldStatus && !newStatus) {
            this._removeValidationError('MultipleMatch');
        }
        this._hasMultipleMatchValidationError = newStatus;
    },
    
    _validateNoMatchError: function() {
        var oldStatus = this._hasNoMatchValidationError;
        var newStatus = false;
        for (var i = 0; i < this._internalSelectedItems.length; i++) {
            if (!Office.Controls.Utils.isNullOrUndefined(this._internalSelectedItems[i]._optionsList) && !this._internalSelectedItems[i]._optionsList.length) {
                newStatus = true;
                break;
            }
        }
        if (!oldStatus && newStatus) {
            this._addValidationError(Office.Controls.PeoplePicker.ValidationError._createNoMatchError$i());
        }
        if (oldStatus && !newStatus) {
            this._removeValidationError('NoMatch');
        }
        this._hasNoMatchValidationError = newStatus;
    },
    
    _displayValidationErrors: function() {
        if (!this._showValidationErrors) {
            return;
        }
        if (!this.errors.length) {
            if (!Office.Controls.Utils.isNullOrUndefined(this._errorMessageElement)) {
                this._errorMessageElement.parentNode.removeChild(this._errorMessageElement);
                this._errorMessageElement = null;
                this._errorDisplayed = null;
            }
        }
        else {
            if (this._errorDisplayed !== this.errors[0]) {
                if (!Office.Controls.Utils.isNullOrUndefined(this._errorMessageElement)) {
                    this._errorMessageElement.parentNode.removeChild(this._errorMessageElement);
                }
                var holderDiv = document.createElement('div');
                holderDiv.innerHTML = Office.Controls._peoplePickerTemplates.generateErrorTemplate(this.errors[0].localizedErrorMessage);
                this._errorMessageElement = holderDiv.firstChild;
                this._root.appendChild(this._errorMessageElement);
                this._errorDisplayed = this.errors[0];
            }
        }
    },
    
    setDataProvider: function(newProvider) {
        this._dataProvider = newProvider;
    }
}


Office.Controls.PeoplePicker._internalPeoplePickerRecord = function(parent, record) {
    this._parent = parent;
    this.Record = record;
}
Office.Controls.PeoplePicker._internalPeoplePickerRecord.prototype = {
    Record: null,
    
    get_record: function() {
        return this.Record;
    },
    
    set_record: function(value) {
        this.Record = value;
        return value;
    },
    
    _principalOptions: null,
    _optionsList: null,
    Node: null,
    
    get_node: function() {
        return this.Node;
    },
    
    set_node: function(value) {
        this.Node = value;
        return value;
    },
    
    _parent: null,

    _onRecordRemovalClick: function(e) {
        var recordRemovalEvent = Office.Controls.Utils.getEvent(e);
        var target = Office.Controls.Utils.getTarget(recordRemovalEvent);
        this._remove();
        Office.Controls.Utils.cancelEvent(e);
        this._parent._autofill.close();
        return false;
    },
    
    _onRecordRemovalKeyDown: function(e) {
        var recordRemovalEvent = Office.Controls.Utils.getEvent(e);
        var target = Office.Controls.Utils.getTarget(recordRemovalEvent);
        if (recordRemovalEvent.keyCode === 8 || recordRemovalEvent.keyCode === 13 || recordRemovalEvent.keyCode === 46) {
            this._remove();
            Office.Controls.Utils.cancelEvent(e);
            this._parent._autofill.close();
        }
        return false;
    },
    
    _add: function() {
        var holderDiv = document.createElement('div');
        holderDiv.innerHTML = Office.Controls._peoplePickerTemplates.generateRecordTemplate(this.Record, this._parent._allowMultiple);
        var recordElement = holderDiv.firstChild;
        var removeButtonElement = recordElement.querySelector('div.ms-PeoplePicker-personaRemove');
        var $$t_5 = this;
        Office.Controls.Utils.addEventListener(removeButtonElement, 'click', function(e) {
            return $$t_5._onRecordRemovalClick(e);
        });
        var $$t_6 = this;
        Office.Controls.Utils.addEventListener(removeButtonElement, 'keydown', function(e) {
            return $$t_6._onRecordRemovalKeyDown(e);
        });
        this._parent._resolvedListRoot.appendChild(recordElement);
        this._parent._defaultText.className = 'office-hide';
        this.Node = recordElement;
    },
    
    _remove: function() {
        this._removeAndNotTriggerUserListener();
        this._parent._onDataRemoved(this.Record);
        this._parent._textInput.focus();
    },
    
    _removeAndNotTriggerUserListener: function() {
        this._parent._resolvedListRoot.removeChild(this.Node);
        for (var i = 0; i < this._parent._internalSelectedItems.length; i++) {
            if (this._parent._internalSelectedItems[i] === this) {
                this._parent._internalSelectedItems.splice(i, 1);
            }
        }
        for (var i = 0; i < this._parent.selectedItems.length; i++) {
            if (this._parent.selectedItems[i] === this.Record) {
                this._parent.selectedItems.splice(i, 1);
            }
        }
    },
    
    _setResolveOptions: function(options) {
        this._optionsList = options;
        this._principalOptions = {};
        for (var i = 0; i < options.length; i++) {
            this._principalOptions[options[i].PersonId] = options[i];
        }
        var $$t_3 = this;
        Office.Controls.Utils.addEventListener(this.Node, 'click', function(e) {
            return $$t_3._onUnresolvedUserClick(e);
        });
        this._parent._validateMultipleMatchError();
        this._parent._validateNoMatchError();
    },
    
    _onUnresolvedUserClick: function(e) {
        e = Office.Controls.Utils.getEvent(e);
        this._parent._autofill.flushContent();
        this._parent._autofill.setServerEntries(this._optionsList);
        var $$t_2 = this;
        this._parent._autofill.open(function(selectedPrincipal) {
            $$t_2._onAutofillClick(selectedPrincipal);
        });
        this._parent._autofill.focusOnFirstElement();
        Office.Controls.Utils.cancelEvent(e);
        return false;
    },
    
    _resolveTo: function(principal) {
        Office.Controls.PeoplePicker._copyToRecord$i(this.Record, principal);
        this.Record.text = principal.DisplayName;
        this.Record.isResolved = true;
        if (this._parent.enableCache) {
            this._parent._addToCache(principal);
        }
        Office.Controls.Utils.removeClass(this.Node, 'has-error');
        var primaryTextNode = this.Node.querySelector('div.ms-Persona-primaryText');
        primaryTextNode.innerHTML = Office.Controls.Utils.htmlEncode(principal.DisplayName);
        this._updateHoverText(primaryTextNode);
    },
    
    _refresh: function(principal) {
        Office.Controls.PeoplePicker._copyToRecord$i(this.Record, principal);
        this.Record.text = principal.DisplayName;
        var primaryTextNode = this.Node.querySelector('div.ms-Persona-primaryText');
        primaryTextNode.innerHTML = Office.Controls.Utils.htmlEncode(principal.DisplayName);
    },
    
    _unresolve: function() {
        this.Record.isResolved = false;
        Office.Controls.Utils.addClass(this.Node, 'has-error');
        var primaryTextNode = this.Node.querySelector('div.ms-Persona-primaryText');
        primaryTextNode.innerHTML = Office.Controls.Utils.htmlEncode(this.Record.text);
        this._updateHoverText(primaryTextNode);
    },
    
    _updateHoverText: function(userLabel) {
        userLabel.title = Office.Controls.Utils.htmlEncode(this.Record.text);
        this.Node.querySelector('div.ms-PeoplePicker-personaRemove').title = Office.Controls.Utils.formatString(Office.Controls._peoplePickerTemplates.getString(Office.Controls.Utils.htmlEncode('PP_RemovePerson')), this.Record.text);
    },
    
    _onAutofillClick: function(selectedPrincipal) {
        this._parent._onRemoved(this._parent, this.Record.info);
        this._resolveTo(selectedPrincipal);
        this._parent._refreshInputField();
        this._principalOptions = null;
        this._optionsList = null;
        if (this._parent.enableCache) {
            this._parent._addToCache(selectedPrincipal);
        }
        this._parent._validateMultipleMatchError();
        this._parent._autofill.close();
        this._parent._onAdded(this._parent, this.Record);
        this._parent._onChange(this._parent);
    }
}


Office.Controls.PeoplePicker._autofillContainer = function(parent) {
    this._entries = {};
    this._cachedEntries = new Array(0);
    this._serverEntries = new Array(0);
    this._parent = parent;
    this._root = parent._autofillElement;
    if (!Office.Controls.PeoplePicker._autofillContainer._boolBodyHandlerAdded) {
        var $$t_2 = this;
        Office.Controls.Utils.addEventListener(document.body, 'click', function(e) {
            return Office.Controls.PeoplePicker._autofillContainer._bodyOnClick(e);
        });
        Office.Controls.PeoplePicker._autofillContainer._boolBodyHandlerAdded = true;
    }
}
Office.Controls.PeoplePicker._autofillContainer._getControlRootFromSubElement = function(element) {
    while (element && element.nodeName.toLowerCase() !== 'body') {
        if (element.className.indexOf('office office-peoplepicker') !== -1) {
            return element;
        }
        element = element.parentNode;
    }
    return null;
}
Office.Controls.PeoplePicker._autofillContainer._bodyOnClick = function(e) {
    if (!Office.Controls.PeoplePicker._autofillContainer.currentOpened) {
        return true;
    }
    var click = Office.Controls.Utils.getEvent(e);
    var target = Office.Controls.Utils.getTarget(click);
    var controlRoot = Office.Controls.PeoplePicker._autofillContainer._getControlRootFromSubElement(target);
    if (!target || controlRoot !== Office.Controls.PeoplePicker._autofillContainer.currentOpened._parent._root) {
        Office.Controls.PeoplePicker._autofillContainer.currentOpened.close();
    }
    return true;
}
Office.Controls.PeoplePicker._autofillContainer.prototype = {
    _parent: null,
    _root: null,
    IsDisplayed: false,
    
    get_isDisplayed: function() {
        return this.IsDisplayed;
    },
    
    set_isDisplayed: function(value) {
        this.IsDisplayed = value;
        return value;
    },
    
    setCachedEntries: function(entries) {
        this._cachedEntries = entries;
        this._entries = {};
        var length = entries.length;
        for (var i = 0; i < length; i++) {
            this._entries[entries[i].PersonId] = entries[i];
        }
    },
    
    getCachedEntries: function() {
        return this._cachedEntries;
    },
    
    getServerEntries: function() {
        return this._serverEntries;
    },
    
    setServerEntries: function(entries) {
        var newServerEntries = new Array(0);
        var length = entries.length;
        for (var i = 0; i < length; i++) {
            var currentEntry = entries[i];
            if (Office.Controls.Utils.isNullOrUndefined(this._entries[currentEntry.PersonId])) {
                this._entries[entries[i].PersonId] = entries[i];
                newServerEntries.push(currentEntry);
            }
        }
        this._serverEntries = newServerEntries;
    },
    
    _renderList: function(handler) {
        var isTabKey = false;
        this._root.innerHTML = Office.Controls._peoplePickerTemplates.generateAutofillListTemplate(this._cachedEntries, this._serverEntries, 30);
        var autofillElementsLinkTags = this._root.querySelectorAll('a');
        for (var i = 0; i < autofillElementsLinkTags.length; i++) {
            var link = autofillElementsLinkTags[i];
            var $$t_A = this;
            Office.Controls.Utils.addEventListener(link, 'click', function(e) {
                return $$t_A._onEntryClick(e, handler);
            });
            var $$t_B = this;
            Office.Controls.Utils.addEventListener(link, 'keydown', function(e) {
                var key = Office.Controls.Utils.getEvent(e);
                isTabKey = (key.keyCode === 9);
                if (key.keyCode === 32 || key.keyCode === 13) {
                    e.preventDefault();
                    e.stopPropagation();
                    return $$t_B._onEntryClick(e, handler);
                }
                return $$t_B._onKeyDown(e);
            });
            var $$t_C = this;
            Office.Controls.Utils.addEventListener(link, 'focus', function(e) {
                return $$t_C._onEntryFocus(e);
            });
            var $$t_D = this;
            Office.Controls.Utils.addEventListener(link, 'blur', function(e) {
                return $$t_D._onEntryBlur(e, isTabKey);
            });
        }
    },
    
    flushContent: function() {
        var entry = this._root.querySelectorAll('div.ms-PeoplePicker-resultGroups');
        for (var i = 0; i < entry.length; i++) {
            this._root.removeChild(entry[i]);
        }
        this._entries = {};
        this._serverEntries = new Array(0);
        this._cachedEntries = new Array(0);
    },
    
    open: function(handler) {
        this._renderList(handler);
        this.IsDisplayed = true;
        Office.Controls.PeoplePicker._autofillContainer.currentOpened = this;
        if (!Office.Controls.Utils.containClass(this._parent._actualRoot, 'is-active')) {
            Office.Controls.Utils.addClass(this._parent._actualRoot, 'is-active');
        }
        if ((this._cachedEntries.length + this._serverEntries.length) > 0) {
            this._parent._changeAlertMessage(Office.Controls._peoplePickerTemplates.getString('PP_SuggestionsAvailable'));
        }
        else {
            this._parent._changeAlertMessage(Office.Controls._peoplePickerTemplates.getString('PP_NoSuggestionsAvailable'));
        }
    },
    
    close: function() {
        this.IsDisplayed = false;
        Office.Controls.Utils.removeClass(this._parent._actualRoot, 'is-active');
    },
    
    openSearchingLoadingStatus: function(searchingName) {
        this._root.innerHTML = Office.Controls._peoplePickerTemplates.generateSerachingLoadingTemplate();
        this.IsDisplayed = true;
        Office.Controls.PeoplePicker._autofillContainer.currentOpened = this;
        if (!Office.Controls.Utils.containClass(this._parent._actualRoot, 'is-active')) {
            Office.Controls.Utils.addClass(this._parent._actualRoot, 'is-active');
        }
    },
    
    closeSearchingLoadingStatus: function() {
        this.IsDisplayed = false;
        Office.Controls.Utils.removeClass(this._parent._actualRoot, 'is-active');
    },
    
    _onEntryClick: function(e, handler) {
        var click = Office.Controls.Utils.getEvent(e);
        var target = Office.Controls.Utils.getTarget(click);
        target = this._getParentListItem(target);
        var PersonId = this._getPersonIdFromListElement(target);
        handler(this._entries[PersonId]);
        this.flushContent();
        return true;
    },
    
    focusOnFirstElement: function() {
        var first = this._root.querySelector('li.ms-PeoplePicker-result');
        if (!Office.Controls.Utils.isNullOrUndefined(first)) {
            first.firstChild.focus();
        }
    },
    
    _onKeyDown: function(e) {
        var key = Office.Controls.Utils.getEvent(e);
        var target = Office.Controls.Utils.getTarget(key);
        if (key.keyCode === 38 || (key.keyCode === 9 && key.shiftKey)) {
            var previous = target.parentNode.previousSibling;
            if (!previous) {
                this._parent._focusToEnd();
            }
            else {
                if (previous.firstChild.tagName.toLowerCase() !== 'a') {
                    previous = previous.previousSibling;
                }
                previous.firstChild.focus();
            }
            Office.Controls.Utils.cancelEvent(e);
            return false;
        }
        else if (key.keyCode === 40) {
            var next = target.parentNode.nextSibling;
            if (next) {
                if (next.firstChild.tagName.toLowerCase() === 'a') {
                    next.firstChild.focus();
                }
                else if (next.firstChild.tagName.toLowerCase() === 'hr' && next.nextSibling && next.nextSibling.firstChild.tagName.toLowerCase() === 'a') {
                    next.nextSibling.firstChild.focus();
                }
            }
        }
        else if (key.keyCode === 27) {
            this.close();
        }
        if (key.keyCode !== 9 && key.keyCode !== 13) {
            Office.Controls.Utils.cancelEvent(key);
        }
        return false;
    },
    
    _getPersonIdFromListElement: function(listElement) {
        return listElement.attributes.getNamedItem('data-office-peoplepicker-value').value;
    },
    
    _getParentListItem: function(element) {
        while (element && element.nodeName.toLowerCase() !== 'li') {
            element = element.parentNode;
        }
        return element;
    },
    
    _onEntryFocus: function(e) {
        var target = Office.Controls.Utils.getTarget(e);
        target = this._getParentListItem(target);
        if (!Office.Controls.Utils.containClass(target, 'office-peoplepicker-autofill-focus')) {
            Office.Controls.Utils.addClass(target, 'office-peoplepicker-autofill-focus');
        }
        return false;
    },
    
    _onEntryBlur: function(e, isTabKey) {
        var target = Office.Controls.Utils.getTarget(e);
        target = this._getParentListItem(target);
        Office.Controls.Utils.removeClass(target, 'office-peoplepicker-autofill-focus');
        if (isTabKey) {
            var next = target.nextSibling;
            if ((next) && (next.nextSibling.className.toLowerCase() === 'ms-PeoplePicker-searchMore js-searchMore'.toLowerCase())) {
                Office.Controls.PeoplePicker._autofillContainer.currentOpened.close();
            }
        }
        return false;
    }
}


















Office.Controls.PeoplePicker.Parameters = function() {}






//Office.Controls.PeoplePicker.ISearchPrincipalDataProvider = function() {}
//Office.Controls.PeoplePicker.ISearchPrincipalDataProvider.registerInterface('Office.Controls.PeoplePicker.ISearchPrincipalDataProvider');


Office.Controls.PeoplePicker._cancelToken = function() {
    this.IsCanceled = false;
}
Office.Controls.PeoplePicker._cancelToken.prototype = {
    IsCanceled: false,
    
    get_isCanceled: function() {
        return this.IsCanceled;
    },
    
    set_isCanceled: function(value) {
        this.IsCanceled = value;
        return value;
    },
    
    cancel: function() {
        this.IsCanceled = true;
    }
}





Office.Controls.PeoplePicker.ValidationError = function() {
}
Office.Controls.PeoplePicker.ValidationError._createMultipleMatchError$i = function() {
    var err = new Office.Controls.PeoplePicker.ValidationError();
    err.errorName = 'MultipleMatch';
    err.localizedErrorMessage = Office.Controls._peoplePickerTemplates.getString('PP_MultipleMatch');
    return err;
}
Office.Controls.PeoplePicker.ValidationError._createMultipleEntryError$i = function() {
    var err = new Office.Controls.PeoplePicker.ValidationError();
    err.errorName = 'MultipleEntry';
    err.localizedErrorMessage = Office.Controls._peoplePickerTemplates.getString('PP_MultipleEntry');
    return err;
}
Office.Controls.PeoplePicker.ValidationError._createNoMatchError$i = function() {
    var err = new Office.Controls.PeoplePicker.ValidationError();
    err.errorName = 'NoMatch';
    err.localizedErrorMessage = Office.Controls._peoplePickerTemplates.getString('PP_NoMatch');
    return err;
}
Office.Controls.PeoplePicker.ValidationError._createServerProblemError$i = function() {
    var err = new Office.Controls.PeoplePicker.ValidationError();
    err.errorName = 'ServerProblem';
    err.localizedErrorMessage = Office.Controls._peoplePickerTemplates.getString('PP_ServerProblem');
    return err;
}
Office.Controls.PeoplePicker.ValidationError.prototype = {
    errorName: null,
    localizedErrorMessage: null
}


Office.Controls.PeoplePicker._mruCache = function() {
    this.isCacheAvailable = this._checkCacheAvailability();
    if (!this.isCacheAvailable) {
        return;
    }
    this._initializeCache();
}
Office.Controls.PeoplePicker._mruCache.getInstance = function() {
    if (!Office.Controls.PeoplePicker._mruCache._instance) {
        Office.Controls.PeoplePicker._mruCache._instance = new Office.Controls.PeoplePicker._mruCache();
    }
    return Office.Controls.PeoplePicker._mruCache._instance;
}
Office.Controls.PeoplePicker._mruCache.prototype = {
    isCacheAvailable: false,
    _localStorage: null,
    _dataObject: null,
    
    get: function(key, maxResults) {
        if (Office.Controls.Utils.isNullOrUndefined(maxResults) || !maxResults) {
            maxResults = 2147483647;
        }
        var numberOfResults = 0;
        var results = new Array(0);
        var cache = this._dataObject.cacheMapping[Office.Controls.Runtime.context.sharePointHostUrl];
        var cacheLength = cache.length;
        for (var i = cacheLength; i > 0 && numberOfResults < maxResults; i--) {
            var candidate = cache[i - 1];
            if (this._entityMatches(candidate, key)) {
                results.push(candidate);
                numberOfResults += 1;
            }
        }
        return results;
    },
    
    set: function(entry) {
        var cache = this._dataObject.cacheMapping[Office.Controls.Runtime.context.sharePointHostUrl];
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
        this._cacheWrite('Office.PeoplePicker.Cache', Office.Controls.Utils.serializeJSON(this._dataObject));
    },
    
    _entityMatches: function(candidate, key) {
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
    
    _initializeCache: function() {
        var cacheData = this._cacheRetreive('Office.PeoplePicker.Cache');
        if (Office.Controls.Utils.isNullOrEmptyString(cacheData)) {
            this._dataObject = new Office.Controls.PeoplePicker._mruCache._mruData();
        }
        else {
            var datas = Office.Controls.Utils.deserializeJSON(cacheData);
            if (datas.cacheVersion) {
                this._dataObject = new Office.Controls.PeoplePicker._mruCache._mruData();
                this._cacheDelete('Office.PeoplePicker.Cache');
            }
            else {
                this._dataObject = datas;
            }
        }
        if (Office.Controls.Utils.isNullOrUndefined(this._dataObject.cacheMapping[Office.Controls.Runtime.context.sharePointHostUrl])) {
            this._dataObject.cacheMapping[Office.Controls.Runtime.context.sharePointHostUrl] = new Array(0);
        }
    },
    
    _checkCacheAvailability: function() {
        this._localStorage = window.self.localStorage;
        if (Office.Controls.Utils.isNullOrUndefined(this._localStorage)) {
            return false;
        }
        return true;
    },
    
    _cacheRetreive: function(key) {
        return this._localStorage.getItem(key);
    },
    
    _cacheWrite: function(key, value) {
        this._localStorage.setItem(key, value);
    },
    
    _cacheDelete: function(key) {
        this._localStorage.removeItem(key);
    }
}


Office.Controls.PeoplePicker._mruCache._mruData = function() {
    this.cacheMapping = {};
    this.cacheVersion = 0;
    this.sharePointHost = Office.Controls.Runtime.context.sharePointHostUrl;
}


Office.Controls.PeoplePickerCustomerInsightStrings = function() {
}


Office.Controls.PeoplePickerResourcesDefaults = function() {
}


Office.Controls._peoplePickerTemplates = function() {
}
Office.Controls._peoplePickerTemplates.getString = function(stringName) {
    var newName = 'PeoplePicker' + stringName.substr(3);
    if ((newName) in Office.Controls.PeoplePicker._res$i) {
        return Office.Controls.PeoplePicker._res$i[newName];
    }
    else {
        return Office.Controls.Utils.getStringFromResource('PeoplePicker', stringName);
    }
}
Office.Controls._peoplePickerTemplates._getDefaultText$i = function(allowMultiple) {
    if (allowMultiple) {
        return Office.Controls._peoplePickerTemplates.getString('PP_DefaultMessagePlural');
    }
    else {
        return Office.Controls._peoplePickerTemplates.getString('PP_DefaultMessage');
    }
}
Office.Controls._peoplePickerTemplates.generateControlTemplate = function(inputName, allowMultiple, inputHint) {
    var defaultText;
    if (Office.Controls.Utils.isNullOrEmptyString(inputHint)) {
        defaultText = Office.Controls.Utils.htmlEncode(Office.Controls._peoplePickerTemplates._getDefaultText$i(allowMultiple));
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
    body += Office.Controls._peoplePickerTemplates.generateAlertNode();
    body += '</div>';
    return body;
}
Office.Controls._peoplePickerTemplates.generateErrorTemplate = function(ErrorMessage) {
    var innerHtml = '<span class=\"office-peoplepicker-error office-error\">';
    innerHtml += Office.Controls.Utils.htmlEncode(ErrorMessage);
    innerHtml += '</span>';
    return innerHtml;
}
Office.Controls._peoplePickerTemplates.generateAutofillListItemTemplate = function(principal, source) {
    var titleText = Office.Controls.Utils.htmlEncode((Office.Controls.Utils.isNullOrEmptyString(principal.Email)) ? '' : principal.Email);
    var itemHtml = '<li tabindex=\"-1\" class=\"ms-PeoplePicker-result\" data-office-peoplepicker-value=\"' + Office.Controls.Utils.htmlEncode(principal.PersonId) + '\" title=\"' + titleText + '\">';
    itemHtml += '<div tabindex=\"-1\" class=\"ms-Persona ms-PersonaAdded\">';
    itemHtml += '<div tabindex=\"-1\" class=\"ms-Persona-details ms-Persona-detailsForDropdownAdded\">';
    itemHtml += '<a onclick=\"return false;\" href=\"#\" tabindex=\"-1\"><div tabindex=\"0\">';
    itemHtml += '<div class=\"ms-Persona-primaryText ms-Persona-primaryTextAdded\" >' + Office.Controls.Utils.htmlEncode(principal.DisplayName) + '</div>';
    if (!Office.Controls.Utils.isNullOrEmptyString(principal.Description)) {
        itemHtml += '<div class=\"ms-Persona-secondaryText ms-Persona-secondaryTextAdded\" >' + Office.Controls.Utils.htmlEncode(principal.Description) + '</div>';
    }
    itemHtml += '</div></a></div></div></li>';
    return itemHtml;
}
Office.Controls._peoplePickerTemplates.generateAutofillListTemplate = function(cachedEntries, serverEntries, maxCount) {
    var html = '<div class=\"ms-PeoplePicker-resultGroups\">';
    if (Office.Controls.Utils.isNullOrUndefined(cachedEntries)) {
        cachedEntries = new Array(0);
    }
    if (Office.Controls.Utils.isNullOrUndefined(serverEntries)) {
        serverEntries = new Array(0);
    }
    html += Office.Controls._peoplePickerTemplates._generateAutofillGroupTemplate(cachedEntries, 1, true);
    html += Office.Controls._peoplePickerTemplates._generateAutofillGroupTemplate(serverEntries, 0, false);
    html += '</div>';
    html += Office.Controls._peoplePickerTemplates.generateAutofillFooterTemplate(cachedEntries.length + serverEntries.length, maxCount);
    return html;
}
Office.Controls._peoplePickerTemplates._generateAutofillGroupTemplate = function(principals, source, isCached) {
    var listHtml = '';
    if (!principals.length) {
        return listHtml;
    }
    var cachedGrouptTitile = Office.Controls.Utils.htmlEncode(Office.Controls._peoplePickerTemplates.getString('PP_SearchResultRecentGroup'));
    var searchedGroupTitile = Office.Controls.Utils.htmlEncode(Office.Controls._peoplePickerTemplates.getString('PP_SearchResultMoreGroup'));
    var groupTitle = (isCached) ? cachedGrouptTitile : searchedGroupTitile;
    listHtml += '<div class=\"ms-PeoplePicker-resultGroup\">';
    listHtml += '<div class=\"ms-PeoplePicker-resultGroupTitle ms-PeoplePicker-resultGroupTitleAdded\">' + groupTitle + '</div>';
    listHtml += '<ul class=\"ms-PeoplePicker-resultList\">';
    for (var i = 0; i < principals.length; i++) {
        listHtml += Office.Controls._peoplePickerTemplates.generateAutofillListItemTemplate(principals[i], source);
    }
    listHtml += '</ul></div>';
    return listHtml;
}
Office.Controls._peoplePickerTemplates.generateAutofillFooterTemplate = function(count, maxCount) {
    var footerHtml = '<div class=\"ms-PeoplePicker-searchMore js-searchMore\">';
    footerHtml += '<div class=\"ms-PeoplePicker-searchMoreIcon\"></div>';
    var footerText;
    if (count >= maxCount) {
        footerText = Office.Controls.Utils.formatString(Office.Controls._peoplePickerTemplates.getString('PP_ShowingTopNumberOfResults'), count.toString());
    }
    else {
        footerText = Office.Controls.Utils.formatString(Office.Controls.Utils.getLocalizedCountValue(Office.Controls._peoplePickerTemplates.getString('PP_Results'), Office.Controls._peoplePickerTemplates.getString('PP_ResultsIntervals'), count), count.toString());
    }
    footerText = Office.Controls.Utils.htmlEncode(footerText);
    footerHtml += '<div class=\"ms-PeoplePicker-searchMorePrimary ms-PeoplePicker-searchMorePrimaryAdded\">' + footerText + '</div>';
    footerHtml += '</div>';
    return footerHtml;
}
Office.Controls._peoplePickerTemplates.generateSerachingLoadingTemplate = function() {
    var searchingLable = Office.Controls.Utils.htmlEncode(Office.Controls._peoplePickerTemplates.getString('PP_Searching'));
    var searchingLoadingHtml = '<div class=\"ms-PeoplePicker-searchMore js-searchMore is-searching\">';
    searchingLoadingHtml += '<div class=\"ms-PeoplePicker-searchMoreIconFixed\"></div>';
    searchingLoadingHtml += '<div class=\"ms-PeoplePicker-searchMorePrimary ms-PeoplePicker-searchMorePrimaryAdded\">' + searchingLable + '</div>';
    searchingLoadingHtml += '</div>';
    return searchingLoadingHtml;
}
Office.Controls._peoplePickerTemplates.generateRecordTemplate = function(record, allowMultiple) {
    var recordHtml;
    var userRecordClass = 'ms-PeoplePicker-persona';
    if (!allowMultiple) {
        userRecordClass += ' ms-PeoplePicker-personaForSingleAdded';
    }
    if (record.isResolved) {
        recordHtml = '<div class=\"' + userRecordClass + '\">';
    }
    else {
        recordHtml = '<div class=\"' + userRecordClass + ' ' + 'has-error' + '\">';
    }
    recordHtml += '<div class=\"ms-Persona ms-Persona--xs ms-PersonaAddedForRecord\">';
    recordHtml += '<div class=\"ms-Persona-details ms-Persona-detailsAdded\">';
    recordHtml += '<div class=\"ms-Persona-primaryText ms-Persona-primaryTextForResolvedUserAdded\">' + Office.Controls.Utils.htmlEncode(record.text);
    recordHtml += '</div></div></div>';
    recordHtml += '<div class=\"ms-PeoplePicker-personaRemove ms-PeoplePicker-personaRemoveAdded\">';
    recordHtml += '<i tabindex=\"0\" class=\"ms-Icon ms-Icon--x ms-Icon-added\">';
    recordHtml += '</i></div>';
    recordHtml += '</div>';
    return recordHtml;
}
Office.Controls._peoplePickerTemplates.generateAlertNode = function() {
    var alertHtml = '<div role=\"alert\" class=\"office-peoplepicker-alert\">';
    alertHtml += '</div>';
    return alertHtml;
}


Office.Controls.Context = function(parameterObject) {
    if (typeof(parameterObject) !== 'object') {
        Office.Controls.Utils.errorConsole('Invalid parameters type');
        return;
    }
    var sharepointHost = parameterObject.sharePointHostUrl;
    if (Office.Controls.Utils.isNullOrUndefined(sharepointHost)) {
        var param = Office.Controls.Utils.getQueryStringParameter('SPHostUrl');
        if (!Office.Controls.Utils.isNullOrEmptyString(param)) {
            param = decodeURIComponent(param);
        }
        this.sharePointHostUrl = param;
    }
    else {
        this.sharePointHostUrl = sharepointHost;
    }
    this.sharePointHostUrl = this.sharePointHostUrl.toLocaleLowerCase();
    var appWeb = parameterObject.appWebUrl;
    if (Office.Controls.Utils.isNullOrUndefined(appWeb)) {
        var param = Office.Controls.Utils.getQueryStringParameter('SPAppWebUrl');
        if (!Office.Controls.Utils.isNullOrEmptyString(param)) {
            param = decodeURIComponent(param);
        }
        this.appWebUrl = param;
    }
    else {
        this.appWebUrl = appWeb;
    }
    this.appWebUrl = this.appWebUrl.toLocaleLowerCase();
    this.requestViaUrl = parameterObject.requestsViaUrl;
}
Office.Controls.Context.prototype = {
    _re: null,
    sharePointHostUrl: null,
    appWebUrl: null,
    requestViaUrl: null,
    
    getRequestExecutor: function() {
        if (!this._re) {
            if (!Office.Controls.Utils.isNullOrEmptyString(Office.Controls.Runtime.context.appWebUrl)) {
                if (!Office.Controls.Utils.isNullOrEmptyString(Office.Controls.Runtime.context.requestViaUrl)) {
                    var options = new SP.RequestExecutorOptions();
                    options.viaUrl = Office.Controls.Runtime.context.requestViaUrl;
                    this._re = new SP.RequestExecutor(Office.Controls.Runtime.context.sharePointHostUrl, options);
                }
                else {
                    this._re = new SP.RequestExecutor(Office.Controls.Runtime.context.appWebUrl);
                }
            }
            else {
                Office.Controls.Utils.errorConsole('Missing authentication informations.');
            }
        }
        return this._re;
    }
}


Office.Controls.Runtime = function() {
}
Office.Controls.Runtime.initialize = function(parameterObject) {
    Office.Controls.Runtime.context = new Office.Controls.Context(parameterObject);
}


Office.Controls.Utils = function() {
}
Office.Controls.Utils.deserializeJSON = function(data) {
    if (Office.Controls.Utils.isNullOrEmptyString(data)) {
        return {};
    }
    else {
        return JSON.parse(data);
    }
}
Office.Controls.Utils.serializeJSON = function(obj) {
    return JSON.stringify(obj);
}
Office.Controls.Utils.isNullOrEmptyString = function(str) {
    var strNull = null;
    return str === strNull || typeof(str) === 'undefined' || !str.length;
}
Office.Controls.Utils.isNullOrUndefined = function(obj) {
    var objNull = null;
    return obj === objNull || typeof(obj) === 'undefined';
}
Office.Controls.Utils.getQueryStringParameter = function(paramToRetrieve) {
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
Office.Controls.Utils.logConsole = function(message) {
    console.log(message);
}
Office.Controls.Utils.warnConsole = function(message) {
    console.warn(message);
}
Office.Controls.Utils.errorConsole = function(message) {
    console.error(message);
}
Office.Controls.Utils._getObjectFromFullyQualifiedName$i = function(objectName) {
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
Office.Controls.Utils.getStringFromResource = function(controlName, stringName) {
    var resourceObjectName = 'Office.Controls.' + controlName + 'Resources';
    var res;
    var nonPreserveCase = stringName.charAt(0).toString().toLowerCase() + stringName.substr(1);
    res = SP.RuntimeRes;
    var str;
    if (!Office.Controls.Utils.isNullOrUndefined(res)) {
        str = res[nonPreserveCase];
        if (!Office.Controls.Utils.isNullOrEmptyString(str)) {
            return str;
        }
    }
    resourceObjectName += 'Defaults';
    res = Office.Controls.Utils._getObjectFromFullyQualifiedName$i(resourceObjectName);
    if (!Office.Controls.Utils.isNullOrUndefined(res)) {
        return res[stringName];
    }
    return stringName;
}
Office.Controls.Utils.addEventListener = function(element, eventName, handler) {
    var h = function(e) {
        try {
            return handler(e);
        }
        catch (ex) {
            //var adapter = Access.ControlTelemetryAdapter.getStaticTelemetryAdapter(null, null);
            //adapter.writeDiagnosticLog('EventListenerException', 'error', ex.message, null);
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
Office.Controls.Utils.getEvent = function(e) {
    return (Office.Controls.Utils.isNullOrUndefined(e)) ? window.event : e;
}
Office.Controls.Utils.getTarget = function(e) {
    return (Office.Controls.Utils.isNullOrUndefined(e.target)) ? e.srcElement : e.target;
}
Office.Controls.Utils.cancelEvent = function(e) {
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
Office.Controls.Utils.addClass = function(elem, className) {
    if (elem.className !== '') {
        elem.className += ' ';
    }
    elem.className += className;
}
Office.Controls.Utils.removeClass = function(elem, className) {
    var regex = new RegExp('( |^)' + className + '( |$)');
    elem.className = elem.className.replace(regex, ' ').trim();
}
Office.Controls.Utils.containClass = function(elem, className) {
    return elem.className.indexOf(className) !== -1;
}
Office.Controls.Utils.cloneData = function(obj) {
    return Office.Controls.Utils.deserializeJSON(Office.Controls.Utils.serializeJSON(obj));
}
Office.Controls.Utils.formatString = function(format) {
    var args = [];
    for (var $ai_8 = 1; $ai_8 < arguments.length; ++$ai_8) {
        args[$ai_8 - 1] = arguments[$ai_8];
    }
    var result = '';
    var i = 0;
    while (i < format.length) {
        var open = Office.Controls.Utils._findPlaceHolder(format, i, '{');
        if (open < 0) {
            result = result + format.substr(i);
            break;
        }
        else {
            var close = Office.Controls.Utils._findPlaceHolder(format, open, '}');
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
Office.Controls.Utils._findPlaceHolder = function(format, start, ch) {
    var index = format.indexOf(ch, start);
    while (index >= 0 && index < format.length - 1 && format.charAt(index + 1) === ch) {
        start = index + 2;
        index = format.indexOf(ch, start);
    }
    return index;
}
Office.Controls.Utils.htmlEncode = function(value) {
    value = value.replace(new RegExp('&', 'g'), '&amp;');
    value = value.replace(new RegExp('\"', 'g'), '&quot;');
    value = value.replace(new RegExp('\'', 'g'), '&#39;');
    value = value.replace(new RegExp('<', 'g'), '&lt;');
    value = value.replace(new RegExp('>', 'g'), '&gt;');
    return value;
}
Office.Controls.Utils.getLocalizedCountValue = function(locText, intervals, count) {
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
Office.Controls.Utils.NOP = function() {
}


Office.Controls.PrincipalInfo.registerClass('Office.Controls.PrincipalInfo');
Office.Controls.PeoplePickerRecord.registerClass('Office.Controls.PeoplePickerRecord');
Office.Controls.PeoplePicker.registerClass('Office.Controls.PeoplePicker');
Office.Controls.PeoplePicker._internalPeoplePickerRecord.registerClass('Office.Controls.PeoplePicker._internalPeoplePickerRecord');
Office.Controls.PeoplePicker._autofillContainer.registerClass('Office.Controls.PeoplePicker._autofillContainer');
Office.Controls.PeoplePicker.Parameters.registerClass('Office.Controls.PeoplePicker.Parameters');
Office.Controls.PeoplePicker._cancelToken.registerClass('Office.Controls.PeoplePicker._cancelToken');
Office.Controls.PeoplePicker.ValidationError.registerClass('Office.Controls.PeoplePicker.ValidationError');
Office.Controls.PeoplePicker._mruCache.registerClass('Office.Controls.PeoplePicker._mruCache');
Office.Controls.PeoplePicker._mruCache._mruData.registerClass('Office.Controls.PeoplePicker._mruCache._mruData');
Office.Controls.PeoplePickerCustomerInsightStrings.registerClass('Office.Controls.PeoplePickerCustomerInsightStrings');
Office.Controls.PeoplePickerResourcesDefaults.registerClass('Office.Controls.PeoplePickerResourcesDefaults');
Office.Controls._peoplePickerTemplates.registerClass('Office.Controls._peoplePickerTemplates');
Office.Controls.Context.registerClass('Office.Controls.Context');
Office.Controls.Runtime.registerClass('Office.Controls.Runtime');
Office.Controls.Utils.registerClass('Office.Controls.Utils');
Office.Controls.PeoplePicker._res$i = null;
Office.Controls.PeoplePicker._autofillContainer.currentOpened = null;
Office.Controls.PeoplePicker._autofillContainer._boolBodyHandlerAdded = false;
Office.Controls.PeoplePicker.ValidationError.multipleMatchName = 'MultipleMatch';
Office.Controls.PeoplePicker.ValidationError.multipleEntryName = 'MultipleEntry';
Office.Controls.PeoplePicker.ValidationError.noMatchName = 'NoMatch';
Office.Controls.PeoplePicker.ValidationError.serverProblemName = 'ServerProblem';
Office.Controls.PeoplePicker._mruCache._instance = null;
Office.Controls.PeoplePickerCustomerInsightStrings.addPeople = 'AddPeople';
Office.Controls.PeoplePickerCustomerInsightStrings.deletePeople = 'DeletePeople';
Office.Controls.PeoplePickerCustomerInsightStrings.selectSearchGrouptAction = 'selectSearchGrouptAction';
Office.Controls.PeoplePickerCustomerInsightStrings.inputBeginAction = 'inputBeginAction';
Office.Controls.PeoplePickerCustomerInsightStrings.selectingFromCache = 'selectingFromCache';
Office.Controls.PeoplePickerCustomerInsightStrings.selectingIndex = 'selectingIndex';
Office.Controls.PeoplePickerCustomerInsightStrings.searchingTimes = 'searchingTimes';
Office.Controls.PeoplePickerCustomerInsightStrings.peoplePickerType = 'peoplePickerType';
Office.Controls.PeoplePickerCustomerInsightStrings.createPeoplePicker = 'createPeoplePicker';
Office.Controls.PeoplePickerCustomerInsightStrings.tryToAddMultiplyPeople = 'TryToAddMultiplyPeople';
Office.Controls.PeoplePickerCustomerInsightStrings.selectPeople = 'SelectPeople';
Office.Controls.PeoplePickerCustomerInsightStrings.principalType = 'PrincipalType';
Office.Controls.PeoplePickerCustomerInsightStrings.inputType = 'InputType';
Office.Controls.PeoplePickerCustomerInsightStrings.inputLength = 'InputLength';
Office.Controls.PeoplePickerCustomerInsightStrings.resultNumber = 'ResultNumber';
Office.Controls.PeoplePickerCustomerInsightStrings.controlType = 'PeoplePicker';
Office.Controls.PeoplePickerCustomerInsightStrings.howInvoke = 'HowInvoke';
Office.Controls.PeoplePickerCustomerInsightStrings.queryCancel = 'QueryCancel';
Office.Controls.PeoplePickerCustomerInsightStrings.queryError = 'QueryError';
Office.Controls.PeoplePickerCustomerInsightStrings.controlInitException = 'ControlInitException';
Office.Controls.PeoplePickerCustomerInsightStrings.PersonId = 'PersonId';
Office.Controls.PeoplePickerCustomerInsightStrings.displayName = 'DisplayName';
Office.Controls.PeoplePickerResourcesDefaults.PP_SuggestionsAvailable = 'Suggestions Available';
Office.Controls.PeoplePickerResourcesDefaults.PP_NoMatch = 'We couldn\'t find an exact match.';
Office.Controls.PeoplePickerResourcesDefaults.PP_ShowingTopNumberOfResults = '{0} found';
Office.Controls.PeoplePickerResourcesDefaults.PP_ServerProblem = 'Sorry, we\'re having trouble reaching the server.';
Office.Controls.PeoplePickerResourcesDefaults.PP_DefaultMessagePlural = 'Enter names or email addresses...';
Office.Controls.PeoplePickerResourcesDefaults.PP_MultipleMatch = 'Multiple entries matched, please click to resolve.';
Office.Controls.PeoplePickerResourcesDefaults.PP_Results = 'No results found||{0} found||Showing {0} results';
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
