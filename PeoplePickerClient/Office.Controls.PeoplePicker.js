/*! Version=16.00.0549.000 */
(function(){

Type.registerNamespace('Office.Controls');

Office.Controls.PrincipalInfo = function() {}


Office.Controls.PeoplePickerRecord = function() {
}
Office.Controls.PeoplePickerRecord.prototype = {
    isResolved: false,
    text: null,
    department: null,
    displayName: null,
    email: null,
    jobTitle: null,
    loginName: null,
    mobile: null,
    principalId: 0,
    principalType: 0,
    sipAddress: null
}


Office.Controls.PeoplePicker = function(root, parameterObject, dataProvider) {
    this._currentTimerId$p$0 = -1;
    this.selectedItems = new Array(0);
    this._internalSelectedItems$p$0 = new Array(0);
    this.errors = new Array(0);
    this._cache$p$0 = Office.Controls.PeoplePicker._mruCache.getInstance();
    this._controlTelemetryAdapter$p$0 = new Access.ControlTelemetryAdapter('PeoplePicker', null);
    try {
        if (typeof(root) !== 'object' || typeof(parameterObject) !== 'object') {
            Office.Controls.Utils.errorConsole('Invalid parameters type');
            return;
        }
        this._root$p$0 = root;
        this._allowMultiple$p$0 = parameterObject.allowMultipleSelections;
        this._groupName$p$0 = parameterObject.groupName;
        Office.Controls.PeoplePicker._res$i = parameterObject.res;
        if (Office.Controls.Utils.isNullOrUndefined(Office.Controls.PeoplePicker._res$i)) {
            Office.Controls.PeoplePicker._res$i = {};
        }
        this._onAdded$p$0 = parameterObject.onAdded;
        if (Office.Controls.Utils.isNullOrUndefined(this._onAdded$p$0)) {
            this._onAdded$p$0 = Office.Controls.PeoplePicker._nopAddRemove$p;
        }
        this._onRemoved$p$0 = parameterObject.onRemoved;
        if (Office.Controls.Utils.isNullOrUndefined(this._onRemoved$p$0)) {
            this._onRemoved$p$0 = Office.Controls.PeoplePicker._nopAddRemove$p;
        }
        this._onChange$p$0 = parameterObject.onChange;
        if (Office.Controls.Utils.isNullOrUndefined(this._onChange$p$0)) {
            this._onChange$p$0 = Office.Controls.PeoplePicker._nopOperation$p;
        }
        this._onFocus$p$0 = parameterObject.onFocus;
        if (Office.Controls.Utils.isNullOrUndefined(this._onFocus$p$0)) {
            this._onFocus$p$0 = Office.Controls.PeoplePicker._nopOperation$p;
        }
        this._onBlur$p$0 = parameterObject.onBlur;
        if (Office.Controls.Utils.isNullOrUndefined(this._onBlur$p$0)) {
            this._onBlur$p$0 = Office.Controls.PeoplePicker._nopOperation$p;
        }
        this._onError$p$0 = parameterObject.onError;
        //if (!dataProvider) {
        //    this._dataProvider$p$0 = new Office.Controls.PeoplePicker._searchPrincipalServerDataProvider();
        //    (this._dataProvider$p$0).setControlTelemetryAdapter(this._controlTelemetryAdapter$p$0);
       // }
       // else {
            this._dataProvider$p$0 = dataProvider;
       // }
        if (Office.Controls.Utils.isNullOrUndefined(parameterObject.displayErrors)) {
            this._showValidationErrors$p$0 = true;
        }
        else {
            this._showValidationErrors$p$0 = parameterObject.displayErrors;
        }
        if (!Office.Controls.Utils.isNullOrEmptyString(parameterObject.placeholder)) {
            this._defaultTextOverride$p$0 = parameterObject.placeholder;
        }
        if (Office.Controls.Utils.isNullOrUndefined(parameterObject.showInputHint)) {
            this._showInputHint$p$0 = true;
        }
        else {
            this._showInputHint$p$0 = parameterObject.showInputHint;
        }
        if (Office.Controls.Utils.isNullOrUndefined(parameterObject.showDistributionGroups)) {
            this._showDistributionGroups$p$0 = true;
        }
        else {
            this._showDistributionGroups$p$0 = parameterObject.showDistributionGroups;
        }
        this._inputTabindex$p$0 = parameterObject.inputTabindex;
        this._renderControl$p$0(parameterObject.inputName);
        this._autofill$p$0 = new Office.Controls.PeoplePicker._autofillContainer(this);
        var loggingProperties = {};
        loggingProperties['peoplePickerType'] = root.getAttribute('name');
        this._controlTelemetryAdapter$p$0.logSingletonCustomerAction('createPeoplePicker', loggingProperties, '');
    }
    catch (ex) {
        this._controlTelemetryAdapter$p$0.writeDiagnosticLog('ControlInitException', 'error', ex.message, null);
        throw ex;
    }
}
Office.Controls.PeoplePicker._copyToRecord$i = function(record, info) {
    record.department = info.Department;
    record.displayName = info.DisplayName;
    record.email = info.Email;
    record.jobTitle = info.JobTitle;
    record.loginName = info.LoginName;
    record.mobile = info.Mobile;
    record.principalId = info.PrincipalId;
    record.principalType = info.PrincipalType;
    record.sipAddress = info.SIPAddress;
}
Office.Controls.PeoplePicker._parseUserPaste$p = function(content) {
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
Office.Controls.PeoplePicker._nopAddRemove$p = function(p1, p2) {
}
Office.Controls.PeoplePicker._nopOperation$p = function(p1) {
}
Office.Controls.PeoplePicker.create = function(root, parameterObject) {
    return new Office.Controls.PeoplePicker(root, parameterObject);
}
Office.Controls.PeoplePicker.prototype = {
    _allowMultiple$p$0: false,
    _groupName$p$0: null,
    _defaultTextOverride$p$0: null,
    _onAdded$p$0: null,
    _onRemoved$p$0: null,
    _onChange$p$0: null,
    _onFocus$p$0: null,
    _onBlur$p$0: null,
    _onError$p$0: null,
    _dataProvider$p$0: null,
    _showValidationErrors$p$0: false,
    _showInputHint$p$0: false,
    _showDistributionGroups$p$0: false,
    _inputTabindex$p$0: 0,
    _searchingTimes$p$0: 0,
    _inputBeginAction$p$0: false,
    _actualRoot$p$0: null,
    _textInput$p$0: null,
    _inputData$p$0: null,
    _defaultText$p$0: null,
    _resolvedListRoot$p$0: null,
    _autofillElement$p$0: null,
    _errorMessageElement$p$0: null,
    _root$p$0: null,
    _alertDiv$p$0: null,
    _lastSearchQuery$p$0: '',
    _currentToken$p$0: null,
    _widthSet$p$0: false,
    _currentPrincipalsChoices$p$0: null,
    hasErrors: false,
    _errorDisplayed$p$0: null,
    _hasMultipleEntryValidationError$p$0: false,
    _hasMultipleMatchValidationError$p$0: false,
    _hasNoMatchValidationError$p$0: false,
    _autofill$p$0: null,
    
    reset: function() {
        while (this._internalSelectedItems$p$0.length) {
            var record = this._internalSelectedItems$p$0[0];
            record._removeAndNotTriggerUserListener$i$0();
        }
        this._setTextInputDisplayStyle$p$0();
        this._validateMultipleMatchError$p$0();
        this._validateMultipleEntryAllowed$p$0();
        this._validateNoMatchError$p$0();
        this._clearInputField$p$0();
        if (Office.Controls.PeoplePicker._autofillContainer.currentOpened) {
            Office.Controls.PeoplePicker._autofillContainer.currentOpened.close();
        }
        this._toggleDefaultText$p$0();
    },
    
    remove: function(entryToRemove) {
        var record = this._internalSelectedItems$p$0;
        for (var i = 0; i < record.length; i++) {
            if (record[i]._$$pf_Record$p$0 === entryToRemove) {
                record[i]._removeAndNotTriggerUserListener$i$0();
                break;
            }
        }
    },
    
    add: function(p1, resolve) {
        if (typeof(p1) === 'string') {
            this._addThroughString$p$0(p1);
        }
        else {
            if (Office.Controls.Utils.isNullOrUndefined(resolve)) {
                this._addThroughRecord$p$0(p1, false);
            }
            else {
                this._addThroughRecord$p$0(p1, resolve);
            }
        }
    },
    
    getUserInfoAsync: function(userInfoHandler, userEmail) {
        var scopes = (this._showDistributionGroups$p$0) ? 15 : 13;
        var record = new Office.Controls.PeoplePickerRecord();
        var $$t_6 = this, $$t_7 = this;
        this._dataProvider$p$0.getPrincipals(userEmail, scopes, 15, null, 1, function(principalsReceived) {
            Office.Controls.PeoplePicker._copyToRecord$i(record, principalsReceived[0]);
            userInfoHandler(record);
        }, function(error) {
            userInfoHandler(null);
        });
    },
    
    get_textInput: function() {
        return this._textInput$p$0;
    },
    
    get_actualRoot: function() {
        return this._actualRoot$p$0;
    },
    
    _addThroughString$p$0: function(input) {
        if (Office.Controls.Utils.isNullOrEmptyString(input)) {
            Office.Controls.Utils.errorConsole('Input can\'t be null or empty string. PeoplePicker Id : ' + this._root$p$0.id);
            return;
        }
        this._addUnresolvedPrincipal$p$0(input, false);
    },
    
    _addThroughRecord$p$0: function(info, resolve) {
        if (resolve) {
            this._addUncertainPrincipal$p$0(info);
        }
        else {
            this._addResolvedRecord$p$0(info);
        }
    },
    
    _renderControl$p$0: function(inputName) {
        this._root$p$0.innerHTML = Office.Controls._peoplePickerTemplates.generateControlTemplate(inputName, this._allowMultiple$p$0, this._defaultTextOverride$p$0);
        if (this._root$p$0.className.length > 0) {
            this._root$p$0.className += ' ';
        }
        this._root$p$0.className += 'office office-peoplepicker';
        this._actualRoot$p$0 = this._root$p$0.querySelector('div.ms-PeoplePicker');
        var $$t_7 = this;
        Office.Controls.Utils.addEventListener(this._actualRoot$p$0, 'click', function(e) {
            return $$t_7._onPickerClick$p$0(e);
        });
        this._inputData$p$0 = this._actualRoot$p$0.querySelector('input[type=\"hidden\"]');
        this._textInput$p$0 = this._actualRoot$p$0.querySelector('input[type=\"text\"]');
        var $$t_8 = this;
        Office.Controls.Utils.addEventListener(this._textInput$p$0, 'focus', function(e) {
            return $$t_8._onInputFocus$p$0(e);
        });
        var $$t_9 = this;
        Office.Controls.Utils.addEventListener(this._textInput$p$0, 'blur', function(e) {
            return $$t_9._onInputBlur$p$0(e);
        });
        var $$t_A = this;
        Office.Controls.Utils.addEventListener(this._textInput$p$0, 'keydown', function(e) {
            return $$t_A._onInputKeyDown$p$0(e);
        });
        var $$t_B = this;
        Office.Controls.Utils.addEventListener(this._textInput$p$0, 'keyup', function(e) {
            return $$t_B._onInputKeyUp$p$0(e);
        });
        var $$t_C = this;
        Office.Controls.Utils.addEventListener(window.self, 'resize', function(e) {
            return $$t_C._onResize$p$0(e);
        });
        this._defaultText$p$0 = this._actualRoot$p$0.querySelector('span.office-peoplepicker-default');
        this._resolvedListRoot$p$0 = this._actualRoot$p$0.querySelector('div.office-peoplepicker-recordList');
        this._autofillElement$p$0 = this._actualRoot$p$0.querySelector('.ms-PeoplePicker-results');
        this._alertDiv$p$0 = this._actualRoot$p$0.querySelector('.office-peoplepicker-alert');
        this._toggleDefaultText$p$0();
        if (!Office.Controls.Utils.isNullOrUndefined(this._inputTabindex$p$0)) {
            this._textInput$p$0.setAttribute('tabindex', this._inputTabindex$p$0);
        }
    },
    
    _toggleDefaultText$p$0: function() {
        if (this._root$p$0.clientWidth > 200 && this._actualRoot$p$0.className.indexOf('office-peoplepicker-autofill-focus') === -1 && this._showInputHint$p$0 && !this.selectedItems.length && !this._textInput$p$0.value.length) {
            this._defaultText$p$0.className = 'office-peoplepicker-default office-helper';
        }
        else {
            this._defaultText$p$0.className = 'office-hide';
        }
    },
    
    _onResize$p$0: function(e) {
        this._toggleDefaultText$p$0();
        return true;
    },
    
    _onInputKeyDown$p$0: function(e) {
        var keyEvent = Office.Controls.Utils.getEvent(e);
        if (keyEvent.keyCode === 27) {
            this._autofill$p$0.close();
        }
        else if (keyEvent.keyCode === 40 && this._autofill$p$0._$$pf_IsDisplayed$p$0) {
            var firstElement = this._autofillElement$p$0.querySelector('a');
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
                range.moveStart('character', -this._textInput$p$0.value.length);
                var caretPos = range.text.length;
                if (!selectedText.length && !caretPos) {
                    shouldRemove = true;
                }
            }
            else {
                var selectionStart = this._textInput$p$0.selectionStart;
                var selectionEnd = this._textInput$p$0.selectionEnd;
                if (!selectionStart && selectionStart === selectionEnd) {
                    shouldRemove = true;
                }
            }
            if (shouldRemove && this._internalSelectedItems$p$0.length) {
                var correlationId = Access.TelemetryManager.generateGuid();
                var loggingProperties = {};
                loggingProperties['HowInvoke'] = keyEvent.keyCode.toString();
                this._controlTelemetryAdapter$p$0.writeCustomerActionLog('DeletePeople', 'start', loggingProperties, correlationId);
                Access.TelemetryManager.get_contextManager().storeCrossScopeCorrelationId(this._controlTelemetryAdapter$p$0.getCorrelationKey('DeletePeople'), correlationId);
                this._internalSelectedItems$p$0[this._internalSelectedItems$p$0.length - 1]._remove$i$0();
            }
        }
        else if ((keyEvent.keyCode === 75 && keyEvent.ctrlKey) || (keyEvent.keyCode === 186) || (keyEvent.keyCode === 9 && this._autofill$p$0._$$pf_IsDisplayed$p$0) || (keyEvent.keyCode === 13)) {
            keyEvent.preventDefault();
            keyEvent.stopPropagation();
            this._cancelLastRequest$p$0();
            this._attemptResolveInput$p$0();
            Office.Controls.Utils.cancelEvent(e);
            return false;
        }
        else if ((keyEvent.keyCode === 86 && keyEvent.ctrlKey) || (keyEvent.keyCode === 186)) {
            this._cancelLastRequest$p$0();
            var $$t_C = this;
            window.setTimeout(function() {
                $$t_C._textInput$p$0.value = Office.Controls.PeoplePicker._parseUserPaste$p($$t_C._textInput$p$0.value);
                $$t_C._attemptResolveInput$p$0();
            }, 0);
            return true;
        }
        else if (keyEvent.keyCode === 13 && keyEvent.shiftKey) {
            var $$t_D = this;
            this._autofill$p$0.open(function(selectedPrincipal) {
                $$t_D._addResolvedPrincipal$p$0(selectedPrincipal);
            });
        }
        else {
            this._resizeInputField$p$0();
        }
        return true;
    },
    
    _cancelLastRequest$p$0: function() {
        window.clearTimeout(this._currentTimerId$p$0);
        if (!Office.Controls.Utils.isNullOrUndefined(this._currentToken$p$0)) {
            this._hideLoadingIcon$p$0();
            this._currentToken$p$0.cancel();
            this._currentToken$p$0 = null;
        }
    },
    
    _onInputKeyUp$p$0: function(e) {
        if (!this._inputBeginAction$p$0) {
            this._controlTelemetryAdapter$p$0.startPerformance('inputBeginAction', null, '');
            this._inputBeginAction$p$0 = true;
        }
        this._startQueryAfterDelay$p$0();
        this._resizeInputField$p$0();
        this._autofill$p$0.close();
        return true;
    },
    
    _displayCachedEntries$p$0: function() {
        var cachedEntries = this._cache$p$0.get(this._textInput$p$0.value, 5);
        this._autofill$p$0.setCachedEntries(cachedEntries);
        if (!cachedEntries.length) {
            return;
        }
        var $$t_2 = this;
        this._autofill$p$0.open(function(selectedPrincipal) {
            $$t_2._addResolvedPrincipal$p$0(selectedPrincipal);
        });
    },
    
    _resizeInputField$p$0: function() {
        var size = Math.max(this._textInput$p$0.value.length + 1, 1);
        this._textInput$p$0.size = size;
    },
    
    _clearInputField$p$0: function() {
        this._textInput$p$0.value = '';
        this._resizeInputField$p$0();
    },
    
    _startQueryAfterDelay$p$0: function() {
        var crossScopeCorrelationId = Access.TelemetryManager.generateGuid();
        var loggingProperties = {};
        this._cancelLastRequest$p$0();
        var currentValue = this._textInput$p$0.value;
        var scopes = (this._showDistributionGroups$p$0) ? 15 : 13;
        var $$t_7 = this;
        this._currentTimerId$p$0 = window.setTimeout(function() {
            if (currentValue !== $$t_7._lastSearchQuery$p$0) {
                $$t_7._lastSearchQuery$p$0 = currentValue;
                if (currentValue.length >= 3) {
                    $$t_7._searchingTimes$p$0++;
                    loggingProperties['InputLength'] = currentValue.length;
                    $$t_7._controlTelemetryAdapter$p$0.writeCustomerActionLog('AddPeople', 'start', loggingProperties, crossScopeCorrelationId);
                    Access.TelemetryManager.get_contextManager().storeCrossScopeCorrelationId($$t_7._controlTelemetryAdapter$p$0.getCorrelationKey('AddPeople'), crossScopeCorrelationId);
                    $$t_7._controlTelemetryAdapter$p$0.startPerformance('AddPeople', null, '');
                    $$t_7._displayLoadingIcon$p$0(currentValue);
                    $$t_7._removeValidationError$p$0('ServerProblem');
                    var token = new Office.Controls.PeoplePicker._cancelToken();
                    $$t_7._currentToken$p$0 = token;
                    $$t_7._dataProvider$p$0.getPrincipals($$t_7._textInput$p$0.value, scopes, 15, $$t_7._groupName$p$0, 30, function(principalsReceived) {
                        loggingProperties = {};
                        if (!token._$$pf_IsCanceled$p$0) {
                            $$t_7._onDataReceived$p$0(principalsReceived);
                            loggingProperties['ResultNumber'] = principalsReceived.length;
                            loggingProperties['InputLength'] = currentValue.length;
                            $$t_7._controlTelemetryAdapter$p$0.writeCustomerActionLog('AddPeople', 'success', loggingProperties, Access.TelemetryManager.get_contextManager().getCrossScopeCorrelationId($$t_7._controlTelemetryAdapter$p$0.getCorrelationKey('AddPeople')));
                            $$t_7._controlTelemetryAdapter$p$0.endPerformance('AddPeople', null, '');
                        }
                        else {
                            $$t_7._hideLoadingIcon$p$0();
                            loggingProperties['HowInvoke'] = 'QueryCancel';
                            $$t_7._controlTelemetryAdapter$p$0.writeCustomerActionLog('AddPeople', 'expectFail', loggingProperties, Access.TelemetryManager.get_contextManager().getCrossScopeCorrelationId($$t_7._controlTelemetryAdapter$p$0.getCorrelationKey('AddPeople')));
                        }
                    }, function(error) {
                        $$t_7._onDataFetchError$p$0(error);
                        loggingProperties = {};
                        loggingProperties['HowInvoke'] = 'QueryError';
                        $$t_7._controlTelemetryAdapter$p$0.writeCustomerActionLog('AddPeople', 'unexpectFail', null, Access.TelemetryManager.get_contextManager().getCrossScopeCorrelationId($$t_7._controlTelemetryAdapter$p$0.getCorrelationKey('AddPeople')));
                        $$t_7._controlTelemetryAdapter$p$0.writeDiagnosticLog(error, 'error', null, null);
                    });
                }
                else {
                    $$t_7._autofill$p$0.close();
                }
                $$t_7._displayCachedEntries$p$0();
            }
        }, 250);
    },
    
    _onDataFetchError$p$0: function(message) {
        this._hideLoadingIcon$p$0();
        this._addValidationError$p$0(Office.Controls.PeoplePicker.ValidationError._createServerProblemError$i());
    },
    
    _onDataReceived$p$0: function(principalsReceived) {
        this._currentPrincipalsChoices$p$0 = {};
        for (var i = 0; i < principalsReceived.length; i++) {
            var principal = principalsReceived[i];
            this._currentPrincipalsChoices$p$0[principal.LoginName] = principal;
        }
        this._autofill$p$0.setServerEntries(principalsReceived);
        this._hideLoadingIcon$p$0();
        var $$t_4 = this;
        this._autofill$p$0.open(function(selectedPrincipal) {
            $$t_4._addResolvedPrincipal$p$0(selectedPrincipal);
        });
    },
    
    _onPickerClick$p$0: function(e) {
        this._textInput$p$0.focus();
        e = Office.Controls.Utils.getEvent(e);
        var element = Office.Controls.Utils.getTarget(e);
        if (element.nodeName.toLowerCase() !== 'input') {
            this._focusToEnd$p$0();
        }
        return true;
    },
    
    _focusToEnd$p$0: function() {
        var endPos = this._textInput$p$0.value.length;
        if (!Office.Controls.Utils.isNullOrUndefined(this._textInput$p$0.createTextRange)) {
            var range = this._textInput$p$0.createTextRange();
            range.collapse(true);
            range.moveStart('character', endPos);
            range.moveEnd('character', endPos);
            range.select();
        }
        else {
            this._textInput$p$0.focus();
            this._textInput$p$0.setSelectionRange(endPos, endPos);
        }
    },
    
    _onInputFocus$p$0: function(e) {
        if (Office.Controls.Utils.isNullOrEmptyString(this._actualRoot$p$0.className)) {
            this._actualRoot$p$0.className = 'office-peoplepicker-autofill-focus';
        }
        else {
            this._actualRoot$p$0.className += ' office-peoplepicker-autofill-focus';
        }
        if (!this._widthSet$p$0) {
            this._setInputMaxWidth$p$0();
        }
        this._toggleDefaultText$p$0();
        this._onFocus$p$0(this);
        return true;
    },
    
    _setInputMaxWidth$p$0: function() {
        var maxwidth = this._actualRoot$p$0.clientWidth - 25;
        if (maxwidth <= 0) {
            maxwidth = 20;
        }
        this._textInput$p$0.style.maxWidth = maxwidth.toString() + 'px';
        this._widthSet$p$0 = true;
    },
    
    _onInputBlur$p$0: function(e) {
        Office.Controls.Utils.removeClass(this._actualRoot$p$0, 'office-peoplepicker-autofill-focus');
        if (this._textInput$p$0.value.length > 0 || this.selectedItems.length > 0) {
            this._onBlur$p$0(this);
            return true;
        }
        this._toggleDefaultText$p$0();
        this._onBlur$p$0(this);
        return true;
    },
    
    _onDataSelected$p$0: function(selectedPrincipal) {
        this._lastSearchQuery$p$0 = '';
        this._validateMultipleEntryAllowed$p$0();
        this._clearInputField$p$0();
        this._refreshInputField$p$0();
        var loggingProperties = {};
        loggingProperties['InputLength'] = this._textInput$p$0.value.length;
        loggingProperties['PrincipalType'] = selectedPrincipal.principalType;
        if (selectedPrincipal.loginName.indexOf(this._textInput$p$0.value) !== -1) {
            loggingProperties['InputType'] = 'LoginName';
        }
        else {
            loggingProperties['InputType'] = 'DisplayName';
        }
        this._controlTelemetryAdapter$p$0.writeInformationalLog('SelectPeople', loggingProperties, null);
    },
    
    _onDataRemoved$p$0: function(selectedPrincipal) {
        this._refreshInputField$p$0();
        this._validateMultipleMatchError$p$0();
        this._validateMultipleEntryAllowed$p$0();
        this._validateNoMatchError$p$0();
        this._onRemoved$p$0(this, selectedPrincipal);
        this._onChange$p$0(this);
        this._controlTelemetryAdapter$p$0.writeCustomerActionLog('DeletePeople', 'success', null, Access.TelemetryManager.get_contextManager().getCrossScopeCorrelationId(this._controlTelemetryAdapter$p$0.getCorrelationKey('DeletePeople')));
    },
    
    _addToCache$p$0: function(entry) {
        if (!this._cache$p$0.isCacheAvailable) {
            return;
        }
        this._cache$p$0.set(entry);
    },
    
    _refreshInputField$p$0: function() {
        this._inputData$p$0.value = Office.Controls.Utils.serializeJSON(this.selectedItems);
        this._setTextInputDisplayStyle$p$0();
    },
    
    _setTextInputDisplayStyle$p$0: function() {
        if ((!this._allowMultiple$p$0) && (this._internalSelectedItems$p$0.length === 1)) {
            this._actualRoot$p$0.className = 'ms-PeoplePicker';
            this._textInput$p$0.className = 'ms-PeoplePicker-searchFieldAddedForSingleSelectionHidden';
            this._textInput$p$0.setAttribute('readonly', 'readonly');
        }
        else {
            this._textInput$p$0.removeAttribute('readonly');
            this._textInput$p$0.className = 'ms-PeoplePicker-searchField ms-PeoplePicker-searchFieldAdded';
        }
    },
    
    _changeAlertMessage$p$0: function(message) {
        this._alertDiv$p$0.innerHTML = Office.Controls.Utils.htmlEncode(message);
    },
    
    _displayLoadingIcon$p$0: function(searchingName) {
        this._changeAlertMessage$p$0(Office.Controls._peoplePickerTemplates.getString('PP_Searching'));
        this._autofill$p$0.openSearchingLoadingStatus(searchingName);
    },
    
    _hideLoadingIcon$p$0: function() {
        this._autofill$p$0.closeSearchingLoadingStatus();
    },
    
    _attemptResolveInput$p$0: function() {
        this._autofill$p$0.close();
        if (this._textInput$p$0.value.length > 0) {
            this._lastSearchQuery$p$0 = '';
            this._addUnresolvedPrincipal$p$0(this._textInput$p$0.value, true);
            this._clearInputField$p$0();
        }
    },
    
    _onDataReceivedForResolve$p$0: function(principalsReceived, internalRecordToResolve) {
        this._hideLoadingIcon$p$0();
        if (principalsReceived.length === 1) {
            internalRecordToResolve._resolveTo$i$0(principalsReceived[0]);
        }
        else {
            internalRecordToResolve._setResolveOptions$i$0(principalsReceived);
        }
        this._refreshInputField$p$0();
        return internalRecordToResolve;
    },
    
    _onDataReceivedForStalenessCheck$p$0: function(principalsReceived, internalRecordToCheck) {
        if (principalsReceived.length === 1) {
            internalRecordToCheck._refresh$i$0(principalsReceived[0]);
        }
        else {
            internalRecordToCheck._unresolve$i$0();
            internalRecordToCheck._setResolveOptions$i$0(principalsReceived);
        }
        this._refreshInputField$p$0();
    },
    
    _addResolvedPrincipal$p$0: function(principal) {
        var record = new Office.Controls.PeoplePickerRecord();
        Office.Controls.PeoplePicker._copyToRecord$i(record, principal);
        record.text = principal.DisplayName;
        record.isResolved = true;
        this.selectedItems.push(record);
        var internalRecord = new Office.Controls.PeoplePicker._internalPeoplePickerRecord(this, record);
        internalRecord.setControlTelemetryAdapter(this._controlTelemetryAdapter$p$0);
        internalRecord._add$i$0();
        this._internalSelectedItems$p$0.push(internalRecord);
        this._onDataSelected$p$0(record);
        this._addToCache$p$0(principal);
        this._currentPrincipalsChoices$p$0 = null;
        this._autofill$p$0.close();
        this._textInput$p$0.focus();
        this._onAdded$p$0(this, record);
        this._onChange$p$0(this);
        var searchingSelectingGrouploggingProperties = {};
        var cacheEntries = this._autofill$p$0.getCachedEntries();
        var selectingFromCache = false;
        var selectingIndex = 0;
        for (var i = 0; i < cacheEntries.length; i++) {
            if (cacheEntries[i].LoginName === principal.LoginName) {
                selectingIndex = i;
                selectingFromCache = true;
                break;
            }
        }
        if (!selectingFromCache) {
            var serverEntries = this._autofill$p$0.getServerEntries();
            for (var i = 0; i < serverEntries.length; i++) {
                if (serverEntries[i].LoginName === principal.LoginName) {
                    selectingIndex = i;
                    break;
                }
            }
        }
        searchingSelectingGrouploggingProperties['searchingTimes'] = this._searchingTimes$p$0;
        searchingSelectingGrouploggingProperties['InputLength'] = this._textInput$p$0.value.length;
        searchingSelectingGrouploggingProperties['selectingFromCache'] = selectingFromCache;
        searchingSelectingGrouploggingProperties['selectingIndex'] = selectingIndex;
        this._controlTelemetryAdapter$p$0.writeInformationalLog('selectSearchGrouptAction', searchingSelectingGrouploggingProperties, '');
        this._searchingTimes$p$0 = 0;
        this._controlTelemetryAdapter$p$0.endPerformance('inputBeginAction', null, '');
        this._inputBeginAction$p$0 = false;
    },
    
    _addResolvedRecord$p$0: function(record) {
        this.selectedItems.push(record);
        var internalRecord = new Office.Controls.PeoplePicker._internalPeoplePickerRecord(this, record);
        internalRecord.setControlTelemetryAdapter(this._controlTelemetryAdapter$p$0);
        internalRecord._add$i$0();
        this._internalSelectedItems$p$0.push(internalRecord);
        this._onDataSelected$p$0(record);
        this._currentPrincipalsChoices$p$0 = null;
    },
    
    _addUncertainPrincipal$p$0: function(record) {
        var scopes = (this._showDistributionGroups$p$0) ? 15 : 13;
        this.selectedItems.push(record);
        var internalRecord = new Office.Controls.PeoplePicker._internalPeoplePickerRecord(this, record);
        internalRecord.setControlTelemetryAdapter(this._controlTelemetryAdapter$p$0);
        internalRecord._add$i$0();
        this._internalSelectedItems$p$0.push(internalRecord);
        var $$t_5 = this, $$t_6 = this;
        this._dataProvider$p$0.getPrincipals(record.email, scopes, 15, this._groupName$p$0, 30, function(ps) {
            $$t_5._onDataReceivedForStalenessCheck$p$0(ps, internalRecord);
        }, function(message) {
            $$t_6._onDataFetchError$p$0(message);
        });
        this._validateMultipleEntryAllowed$p$0();
    },
    
    _addUnresolvedPrincipal$p$0: function(input, triggerUserListener) {
        var scopes = (this._showDistributionGroups$p$0) ? 15 : 13;
        var record = new Office.Controls.PeoplePickerRecord();
        record.text = input;
        record.isResolved = false;
        var internalRecord = new Office.Controls.PeoplePicker._internalPeoplePickerRecord(this, record);
        internalRecord.setControlTelemetryAdapter(this._controlTelemetryAdapter$p$0);
        internalRecord._add$i$0();
        this.selectedItems.push(record);
        this._internalSelectedItems$p$0.push(internalRecord);
        this._displayLoadingIcon$p$0(input);
        var $$t_7 = this, $$t_8 = this;
        this._dataProvider$p$0.getPrincipals(input, scopes, 15, this._groupName$p$0, 30, function(ps) {
            internalRecord = $$t_7._onDataReceivedForResolve$p$0(ps, internalRecord);
            if (triggerUserListener) {
                $$t_7._onAdded$p$0($$t_7, internalRecord._$$pf_Record$p$0);
                $$t_7._onChange$p$0($$t_7);
                if (!Office.Controls.Utils.isNullOrUndefined($$t_7._onError$p$0)) {
                    $$t_7._onError$p$0($$t_7);
                }
            }
        }, function(message) {
            $$t_8._onDataFetchError$p$0(message);
        });
        this._validateMultipleEntryAllowed$p$0();
    },
    
    _addValidationError$p$0: function(err) {
        this.hasErrors = true;
        this.errors.push(err);
        if (!Office.Controls.Utils.isNullOrUndefined(this._onError$p$0)) {
            this._onError$p$0(this);
        }
        else {
            this._displayValidationErrors$p$0();
        }
    },
    
    _removeValidationError$p$0: function(errorName) {
        for (var i = 0; i < this.errors.length; i++) {
            if (this.errors[i].errorName === errorName) {
                this.errors.splice(i, 1);
                break;
            }
        }
        if (!this.errors.length) {
            this.hasErrors = false;
        }
        if (!Office.Controls.Utils.isNullOrUndefined(this._onError$p$0)) {
            this._onError$p$0(this);
        }
        else {
            this._displayValidationErrors$p$0();
        }
    },
    
    _validateMultipleEntryAllowed$p$0: function() {
        if (!this._allowMultiple$p$0) {
            if (this.selectedItems.length > 1) {
                if (!this._hasMultipleEntryValidationError$p$0) {
                    this._addValidationError$p$0(Office.Controls.PeoplePicker.ValidationError._createMultipleEntryError$i());
                    this._hasMultipleEntryValidationError$p$0 = true;
                }
            }
            else if (this._hasMultipleEntryValidationError$p$0) {
                this._removeValidationError$p$0('MultipleEntry');
                this._hasMultipleEntryValidationError$p$0 = false;
            }
        }
    },
    
    _validateMultipleMatchError$p$0: function() {
        var oldStatus = this._hasMultipleMatchValidationError$p$0;
        var newStatus = false;
        for (var i = 0; i < this._internalSelectedItems$p$0.length; i++) {
            if (!Office.Controls.Utils.isNullOrUndefined(this._internalSelectedItems$p$0[i]._optionsList$i$0) && this._internalSelectedItems$p$0[i]._optionsList$i$0.length > 0) {
                newStatus = true;
                break;
            }
        }
        if (!oldStatus && newStatus) {
            this._addValidationError$p$0(Office.Controls.PeoplePicker.ValidationError._createMultipleMatchError$i());
        }
        if (oldStatus && !newStatus) {
            this._removeValidationError$p$0('MultipleMatch');
        }
        this._hasMultipleMatchValidationError$p$0 = newStatus;
    },
    
    _validateNoMatchError$p$0: function() {
        var oldStatus = this._hasNoMatchValidationError$p$0;
        var newStatus = false;
        for (var i = 0; i < this._internalSelectedItems$p$0.length; i++) {
            if (!Office.Controls.Utils.isNullOrUndefined(this._internalSelectedItems$p$0[i]._optionsList$i$0) && !this._internalSelectedItems$p$0[i]._optionsList$i$0.length) {
                newStatus = true;
                break;
            }
        }
        if (!oldStatus && newStatus) {
            this._addValidationError$p$0(Office.Controls.PeoplePicker.ValidationError._createNoMatchError$i());
        }
        if (oldStatus && !newStatus) {
            this._removeValidationError$p$0('NoMatch');
        }
        this._hasNoMatchValidationError$p$0 = newStatus;
    },
    
    _displayValidationErrors$p$0: function() {
        if (!this._showValidationErrors$p$0) {
            return;
        }
        if (!this.errors.length) {
            if (!Office.Controls.Utils.isNullOrUndefined(this._errorMessageElement$p$0)) {
                this._errorMessageElement$p$0.parentNode.removeChild(this._errorMessageElement$p$0);
                this._errorMessageElement$p$0 = null;
                this._errorDisplayed$p$0 = null;
            }
        }
        else {
            if (this._errorDisplayed$p$0 !== this.errors[0]) {
                if (!Office.Controls.Utils.isNullOrUndefined(this._errorMessageElement$p$0)) {
                    this._errorMessageElement$p$0.parentNode.removeChild(this._errorMessageElement$p$0);
                }
                var holderDiv = document.createElement('div');
                holderDiv.innerHTML = Office.Controls._peoplePickerTemplates.generateErrorTemplate(this.errors[0].localizedErrorMessage);
                this._errorMessageElement$p$0 = holderDiv.firstChild;
                this._root$p$0.appendChild(this._errorMessageElement$p$0);
                this._errorDisplayed$p$0 = this.errors[0];
            }
        }
    },
    
    setDataProvider: function(newProvider) {
        this._dataProvider$p$0 = newProvider;
    }
}


Office.Controls.PeoplePicker._internalPeoplePickerRecord = function(parent, record) {
    this._parent$i$0 = parent;
    this._$$pf_Record$p$0 = record;
}
Office.Controls.PeoplePicker._internalPeoplePickerRecord.prototype = {
    _$$pf_Record$p$0: null,
    
    get_record: function() {
        return this._$$pf_Record$p$0;
    },
    
    set_record: function(value) {
        this._$$pf_Record$p$0 = value;
        return value;
    },
    
    _principalOptions$i$0: null,
    _optionsList$i$0: null,
    _$$pf_Node$p$0: null,
    
    get_node: function() {
        return this._$$pf_Node$p$0;
    },
    
    set_node: function(value) {
        this._$$pf_Node$p$0 = value;
        return value;
    },
    
    _parent$i$0: null,
    _adapter$p$0: null,
    
    setControlTelemetryAdapter: function(peoplePickerAdapter) {
        this._adapter$p$0 = peoplePickerAdapter;
    },
    
    _onRecordRemovalClick$p$0: function(e) {
        var recordRemovalEvent = Office.Controls.Utils.getEvent(e);
        var target = Office.Controls.Utils.getTarget(recordRemovalEvent);
        var correlationId = Access.TelemetryManager.generateGuid();
        var loggingProperties = {};
        loggingProperties['HowInvoke'] = 'ByMouseClick';
        this._adapter$p$0.writeCustomerActionLog('DeletePeople', 'start', loggingProperties, correlationId);
        Access.TelemetryManager.get_contextManager().storeCrossScopeCorrelationId(this._adapter$p$0.getCorrelationKey('DeletePeople'), correlationId);
        this._remove$i$0();
        Office.Controls.Utils.cancelEvent(e);
        this._parent$i$0._autofill$p$0.close();
        return false;
    },
    
    _onRecordRemovalKeyDown$p$0: function(e) {
        var recordRemovalEvent = Office.Controls.Utils.getEvent(e);
        var target = Office.Controls.Utils.getTarget(recordRemovalEvent);
        if (recordRemovalEvent.keyCode === 8 || recordRemovalEvent.keyCode === 13 || recordRemovalEvent.keyCode === 46) {
            var correlationId = Access.TelemetryManager.generateGuid();
            var loggingProperties = {};
            loggingProperties['HowInvoke'] = recordRemovalEvent.keyCode.toString();
            this._adapter$p$0.writeCustomerActionLog('DeletePeople', 'start', loggingProperties, correlationId);
            Access.TelemetryManager.get_contextManager().storeCrossScopeCorrelationId(this._adapter$p$0.getCorrelationKey('DeletePeople'), correlationId);
            this._remove$i$0();
            Office.Controls.Utils.cancelEvent(e);
            this._parent$i$0._autofill$p$0.close();
        }
        return false;
    },
    
    _add$i$0: function() {
        var holderDiv = document.createElement('div');
        holderDiv.innerHTML = Office.Controls._peoplePickerTemplates.generateRecordTemplate(this._$$pf_Record$p$0, this._parent$i$0._allowMultiple$p$0);
        var recordElement = holderDiv.firstChild;
        var removeButtonElement = recordElement.querySelector('div.ms-PeoplePicker-personaRemove');
        var $$t_5 = this;
        Office.Controls.Utils.addEventListener(removeButtonElement, 'click', function(e) {
            return $$t_5._onRecordRemovalClick$p$0(e);
        });
        var $$t_6 = this;
        Office.Controls.Utils.addEventListener(removeButtonElement, 'keydown', function(e) {
            return $$t_6._onRecordRemovalKeyDown$p$0(e);
        });
        this._parent$i$0._resolvedListRoot$p$0.appendChild(recordElement);
        this._parent$i$0._defaultText$p$0.className = 'office-hide';
        this._$$pf_Node$p$0 = recordElement;
    },
    
    _remove$i$0: function() {
        this._removeAndNotTriggerUserListener$i$0();
        this._parent$i$0._textInput$p$0.focus();
        this._parent$i$0._onDataRemoved$p$0(this._$$pf_Record$p$0);
    },
    
    _removeAndNotTriggerUserListener$i$0: function() {
        this._parent$i$0._resolvedListRoot$p$0.removeChild(this._$$pf_Node$p$0);
        for (var i = 0; i < this._parent$i$0._internalSelectedItems$p$0.length; i++) {
            if (this._parent$i$0._internalSelectedItems$p$0[i] === this) {
                this._parent$i$0._internalSelectedItems$p$0.splice(i, 1);
            }
        }
        for (var i = 0; i < this._parent$i$0.selectedItems.length; i++) {
            if (this._parent$i$0.selectedItems[i] === this._$$pf_Record$p$0) {
                this._parent$i$0.selectedItems.splice(i, 1);
            }
        }
    },
    
    _setResolveOptions$i$0: function(options) {
        this._optionsList$i$0 = options;
        this._principalOptions$i$0 = {};
        for (var i = 0; i < options.length; i++) {
            this._principalOptions$i$0[options[i].LoginName] = options[i];
        }
        var $$t_3 = this;
        Office.Controls.Utils.addEventListener(this._$$pf_Node$p$0, 'click', function(e) {
            return $$t_3._onUnresolvedUserClick$i$0(e);
        });
        this._parent$i$0._validateMultipleMatchError$p$0();
        this._parent$i$0._validateNoMatchError$p$0();
    },
    
    _onUnresolvedUserClick$i$0: function(e) {
        e = Office.Controls.Utils.getEvent(e);
        this._parent$i$0._autofill$p$0.flushContent();
        this._parent$i$0._autofill$p$0.setServerEntries(this._optionsList$i$0);
        var $$t_2 = this;
        this._parent$i$0._autofill$p$0.open(function(selectedPrincipal) {
            $$t_2._onAutofillClick$i$0(selectedPrincipal);
        });
        this._parent$i$0._autofill$p$0.focusOnFirstElement();
        Office.Controls.Utils.cancelEvent(e);
        return false;
    },
    
    _resolveTo$i$0: function(principal) {
        Office.Controls.PeoplePicker._copyToRecord$i(this._$$pf_Record$p$0, principal);
        this._$$pf_Record$p$0.text = principal.DisplayName;
        this._$$pf_Record$p$0.isResolved = true;
        this._parent$i$0._addToCache$p$0(principal);
        Office.Controls.Utils.removeClass(this._$$pf_Node$p$0, 'has-error');
        var primaryTextNode = this._$$pf_Node$p$0.querySelector('div.ms-Persona-primaryText');
        primaryTextNode.innerHTML = Office.Controls.Utils.htmlEncode(principal.DisplayName);
        this._updateHoverText$p$0(primaryTextNode);
    },
    
    _refresh$i$0: function(principal) {
        Office.Controls.PeoplePicker._copyToRecord$i(this._$$pf_Record$p$0, principal);
        this._$$pf_Record$p$0.text = principal.DisplayName;
        var primaryTextNode = this._$$pf_Node$p$0.querySelector('div.ms-Persona-primaryText');
        primaryTextNode.innerHTML = Office.Controls.Utils.htmlEncode(principal.DisplayName);
    },
    
    _unresolve$i$0: function() {
        this._$$pf_Record$p$0.isResolved = false;
        Office.Controls.Utils.addClass(this._$$pf_Node$p$0, 'has-error');
        var primaryTextNode = this._$$pf_Node$p$0.querySelector('div.ms-Persona-primaryText');
        primaryTextNode.innerHTML = Office.Controls.Utils.htmlEncode(this._$$pf_Record$p$0.text);
        this._updateHoverText$p$0(primaryTextNode);
    },
    
    _updateHoverText$p$0: function(userLabel) {
        userLabel.title = Office.Controls.Utils.htmlEncode(this._$$pf_Record$p$0.text);
        this._$$pf_Node$p$0.querySelector('div.ms-PeoplePicker-personaRemove').title = Office.Controls.Utils.formatString(Office.Controls._peoplePickerTemplates.getString(Office.Controls.Utils.htmlEncode('PP_RemovePerson')), this._$$pf_Record$p$0.text);
    },
    
    _onAutofillClick$i$0: function(selectedPrincipal) {
        this._parent$i$0._onRemoved$p$0(this._parent$i$0, this._$$pf_Record$p$0);
        this._resolveTo$i$0(selectedPrincipal);
        this._parent$i$0._refreshInputField$p$0();
        this._principalOptions$i$0 = null;
        this._optionsList$i$0 = null;
        this._parent$i$0._addToCache$p$0(selectedPrincipal);
        this._parent$i$0._validateMultipleMatchError$p$0();
        this._parent$i$0._autofill$p$0.close();
        this._parent$i$0._onAdded$p$0(this._parent$i$0, this._$$pf_Record$p$0);
        this._parent$i$0._onChange$p$0(this._parent$i$0);
    }
}


Office.Controls.PeoplePicker._autofillContainer = function(parent) {
    this._entries$p$0 = {};
    this._cachedEntries$p$0 = new Array(0);
    this._serverEntries$p$0 = new Array(0);
    this._parent$p$0 = parent;
    this._root$p$0 = parent._autofillElement$p$0;
    if (!Office.Controls.PeoplePicker._autofillContainer._boolBodyHandlerAdded$p) {
        var $$t_2 = this;
        Office.Controls.Utils.addEventListener(document.body, 'click', function(e) {
            return Office.Controls.PeoplePicker._autofillContainer._bodyOnClick$p(e);
        });
        Office.Controls.PeoplePicker._autofillContainer._boolBodyHandlerAdded$p = true;
    }
}
Office.Controls.PeoplePicker._autofillContainer._getControlRootFromSubElement$p = function(element) {
    while (element && element.nodeName.toLowerCase() !== 'body') {
        if (element.className.indexOf('office office-peoplepicker') !== -1) {
            return element;
        }
        element = element.parentNode;
    }
    return null;
}
Office.Controls.PeoplePicker._autofillContainer._bodyOnClick$p = function(e) {
    if (!Office.Controls.PeoplePicker._autofillContainer.currentOpened) {
        return true;
    }
    var click = Office.Controls.Utils.getEvent(e);
    var target = Office.Controls.Utils.getTarget(click);
    var controlRoot = Office.Controls.PeoplePicker._autofillContainer._getControlRootFromSubElement$p(target);
    if (!target || controlRoot !== Office.Controls.PeoplePicker._autofillContainer.currentOpened._parent$p$0._root$p$0) {
        Office.Controls.PeoplePicker._autofillContainer.currentOpened.close();
    }
    return true;
}
Office.Controls.PeoplePicker._autofillContainer.prototype = {
    _parent$p$0: null,
    _root$p$0: null,
    _$$pf_IsDisplayed$p$0: false,
    
    get_isDisplayed: function() {
        return this._$$pf_IsDisplayed$p$0;
    },
    
    set_isDisplayed: function(value) {
        this._$$pf_IsDisplayed$p$0 = value;
        return value;
    },
    
    setCachedEntries: function(entries) {
        this._cachedEntries$p$0 = entries;
        this._entries$p$0 = {};
        var length = entries.length;
        for (var i = 0; i < length; i++) {
            this._entries$p$0[entries[i].LoginName] = entries[i];
        }
    },
    
    getCachedEntries: function() {
        return this._cachedEntries$p$0;
    },
    
    getServerEntries: function() {
        return this._serverEntries$p$0;
    },
    
    setServerEntries: function(entries) {
        var newServerEntries = new Array(0);
        var length = entries.length;
        for (var i = 0; i < length; i++) {
            var currentEntry = entries[i];
            if (Office.Controls.Utils.isNullOrUndefined(this._entries$p$0[currentEntry.LoginName])) {
                this._entries$p$0[entries[i].LoginName] = entries[i];
                newServerEntries.push(currentEntry);
            }
        }
        this._serverEntries$p$0 = newServerEntries;
    },
    
    _renderList$p$0: function(handler) {
        var isTabKey = false;
        this._root$p$0.innerHTML = Office.Controls._peoplePickerTemplates.generateAutofillListTemplate(this._cachedEntries$p$0, this._serverEntries$p$0, 30);
        var autofillElementsLinkTags = this._root$p$0.querySelectorAll('a');
        for (var i = 0; i < autofillElementsLinkTags.length; i++) {
            var link = autofillElementsLinkTags[i];
            var $$t_A = this;
            Office.Controls.Utils.addEventListener(link, 'click', function(e) {
                return $$t_A._onEntryClick$p$0(e, handler);
            });
            var $$t_B = this;
            Office.Controls.Utils.addEventListener(link, 'keydown', function(e) {
                var key = Office.Controls.Utils.getEvent(e);
                isTabKey = (key.keyCode === 9);
                if (key.keyCode === 32 || key.keyCode === 13) {
                    e.preventDefault();
                    e.stopPropagation();
                    return $$t_B._onEntryClick$p$0(e, handler);
                }
                return $$t_B._onKeyDown$p$0(e);
            });
            var $$t_C = this;
            Office.Controls.Utils.addEventListener(link, 'focus', function(e) {
                return $$t_C._onEntryFocus$p$0(e);
            });
            var $$t_D = this;
            Office.Controls.Utils.addEventListener(link, 'blur', function(e) {
                return $$t_D._onEntryBlur$p$0(e, isTabKey);
            });
        }
    },
    
    flushContent: function() {
        var entry = this._root$p$0.querySelectorAll('div.ms-PeoplePicker-resultGroups');
        for (var i = 0; i < entry.length; i++) {
            this._root$p$0.removeChild(entry[i]);
        }
        this._entries$p$0 = {};
        this._serverEntries$p$0 = new Array(0);
        this._cachedEntries$p$0 = new Array(0);
    },
    
    open: function(handler) {
        this._renderList$p$0(handler);
        this._$$pf_IsDisplayed$p$0 = true;
        Office.Controls.PeoplePicker._autofillContainer.currentOpened = this;
        if (!Office.Controls.Utils.containClass(this._parent$p$0._actualRoot$p$0, 'is-active')) {
            Office.Controls.Utils.addClass(this._parent$p$0._actualRoot$p$0, 'is-active');
        }
        if ((this._cachedEntries$p$0.length + this._serverEntries$p$0.length) > 0) {
            this._parent$p$0._changeAlertMessage$p$0(Office.Controls._peoplePickerTemplates.getString('PP_SuggestionsAvailable'));
        }
        else {
            this._parent$p$0._changeAlertMessage$p$0(Office.Controls._peoplePickerTemplates.getString('PP_NoSuggestionsAvailable'));
        }
    },
    
    close: function() {
        this._$$pf_IsDisplayed$p$0 = false;
        Office.Controls.Utils.removeClass(this._parent$p$0._actualRoot$p$0, 'is-active');
    },
    
    openSearchingLoadingStatus: function(searchingName) {
        this._root$p$0.innerHTML = Office.Controls._peoplePickerTemplates.generateSerachingLoadingTemplate();
        this._$$pf_IsDisplayed$p$0 = true;
        Office.Controls.PeoplePicker._autofillContainer.currentOpened = this;
        if (!Office.Controls.Utils.containClass(this._parent$p$0._actualRoot$p$0, 'is-active')) {
            Office.Controls.Utils.addClass(this._parent$p$0._actualRoot$p$0, 'is-active');
        }
    },
    
    closeSearchingLoadingStatus: function() {
        this._$$pf_IsDisplayed$p$0 = false;
        Office.Controls.Utils.removeClass(this._parent$p$0._actualRoot$p$0, 'is-active');
    },
    
    _onEntryClick$p$0: function(e, handler) {
        var click = Office.Controls.Utils.getEvent(e);
        var target = Office.Controls.Utils.getTarget(click);
        target = this._getParentListItem$p$0(target);
        var loginName = this._getLoginNameFromListElement$p$0(target);
        handler(this._entries$p$0[loginName]);
        this.flushContent();
        return true;
    },
    
    focusOnFirstElement: function() {
        var first = this._root$p$0.querySelector('li.ms-PeoplePicker-result');
        if (!Office.Controls.Utils.isNullOrUndefined(first)) {
            first.firstChild.focus();
        }
    },
    
    _onKeyDown$p$0: function(e) {
        var key = Office.Controls.Utils.getEvent(e);
        var target = Office.Controls.Utils.getTarget(key);
        if (key.keyCode === 38 || (key.keyCode === 9 && key.shiftKey)) {
            var previous = target.parentNode.previousSibling;
            if (!previous) {
                this._parent$p$0._focusToEnd$p$0();
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
    
    _getLoginNameFromListElement$p$0: function(listElement) {
        return listElement.attributes.getNamedItem('data-office-peoplepicker-value').value;
    },
    
    _getParentListItem$p$0: function(element) {
        while (element && element.nodeName.toLowerCase() !== 'li') {
            element = element.parentNode;
        }
        return element;
    },
    
    _onEntryFocus$p$0: function(e) {
        var target = Office.Controls.Utils.getTarget(e);
        target = this._getParentListItem$p$0(target);
        if (!Office.Controls.Utils.containClass(target, 'office-peoplepicker-autofill-focus')) {
            Office.Controls.Utils.addClass(target, 'office-peoplepicker-autofill-focus');
        }
        return false;
    },
    
    _onEntryBlur$p$0: function(e, isTabKey) {
        var target = Office.Controls.Utils.getTarget(e);
        target = this._getParentListItem$p$0(target);
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
    this._$$pf_IsCanceled$p$0 = false;
}
Office.Controls.PeoplePicker._cancelToken.prototype = {
    _$$pf_IsCanceled$p$0: false,
    
    get_isCanceled: function() {
        return this._$$pf_IsCanceled$p$0;
    },
    
    set_isCanceled: function(value) {
        this._$$pf_IsCanceled$p$0 = value;
        return value;
    },
    
    cancel: function() {
        this._$$pf_IsCanceled$p$0 = true;
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
    var adapter = Access.ControlTelemetryAdapter.getStaticTelemetryAdapter('PeoplePicker', null);
    adapter.writeInformationalLog('TryToAddMultiplyPeople', null, '');
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
    this.isCacheAvailable = this._checkCacheAvailability$p$0();
    if (!this.isCacheAvailable) {
        return;
    }
    this._initializeCache$p$0();
}
Office.Controls.PeoplePicker._mruCache.getInstance = function() {
    if (!Office.Controls.PeoplePicker._mruCache._instance$p) {
        Office.Controls.PeoplePicker._mruCache._instance$p = new Office.Controls.PeoplePicker._mruCache();
    }
    return Office.Controls.PeoplePicker._mruCache._instance$p;
}
Office.Controls.PeoplePicker._mruCache.prototype = {
    isCacheAvailable: false,
    _localStorage$p$0: null,
    _dataObject$p$0: null,
    
    get: function(key, maxResults) {
        if (Office.Controls.Utils.isNullOrUndefined(maxResults) || !maxResults) {
            maxResults = 2147483647;
        }
        var numberOfResults = 0;
        var results = new Array(0);
        var cache = this._dataObject$p$0.cacheMapping[Office.Controls.Runtime.context.sharePointHostUrl];
        var cacheLength = cache.length;
        for (var i = cacheLength; i > 0 && numberOfResults < maxResults; i--) {
            var candidate = cache[i - 1];
            if (this._entityMatches$p$0(candidate, key)) {
                results.push(candidate);
                numberOfResults += 1;
            }
        }
        return results;
    },
    
    set: function(entry) {
        var cache = this._dataObject$p$0.cacheMapping[Office.Controls.Runtime.context.sharePointHostUrl];
        var cacheSize = cache.length;
        var alreadyThere = false;
        for (var i = 0; i < cacheSize; i++) {
            var cacheEntry = cache[i];
            if (cacheEntry.LoginName === entry.LoginName) {
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
        this._cacheWrite$p$0('Office.PeoplePicker.Cache', Office.Controls.Utils.serializeJSON(this._dataObject$p$0));
    },
    
    _entityMatches$p$0: function(candidate, key) {
        if (Office.Controls.Utils.isNullOrEmptyString(key) || Office.Controls.Utils.isNullOrUndefined(candidate)) {
            return false;
        }
        key = key.toLowerCase();
        var userNameKey = candidate.LoginName;
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
    
    _initializeCache$p$0: function() {
        var cacheData = this._cacheRetreive$p$0('Office.PeoplePicker.Cache');
        if (Office.Controls.Utils.isNullOrEmptyString(cacheData)) {
            this._dataObject$p$0 = new Office.Controls.PeoplePicker._mruCache._mruData();
        }
        else {
            var datas = Office.Controls.Utils.deserializeJSON(cacheData);
            if (datas.cacheVersion) {
                this._dataObject$p$0 = new Office.Controls.PeoplePicker._mruCache._mruData();
                this._cacheDelete$p$0('Office.PeoplePicker.Cache');
            }
            else {
                this._dataObject$p$0 = datas;
            }
        }
        if (Office.Controls.Utils.isNullOrUndefined(this._dataObject$p$0.cacheMapping[Office.Controls.Runtime.context.sharePointHostUrl])) {
            this._dataObject$p$0.cacheMapping[Office.Controls.Runtime.context.sharePointHostUrl] = new Array(0);
        }
    },
    
    _checkCacheAvailability$p$0: function() {
        this._localStorage$p$0 = window.self.localStorage;
        if (Office.Controls.Utils.isNullOrUndefined(this._localStorage$p$0)) {
            return false;
        }
        return true;
    },
    
    _cacheRetreive$p$0: function(key) {
        return this._localStorage$p$0.getItem(key);
    },
    
    _cacheWrite$p$0: function(key, value) {
        this._localStorage$p$0.setItem(key, value);
    },
    
    _cacheDelete$p$0: function(key) {
        this._localStorage$p$0.removeItem(key);
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
Office.Controls._peoplePickerTemplates.generateControlTemplate = function(inputName, allowMultiple, defaultTextOverride) {
    var defaultText;
    if (Office.Controls.Utils.isNullOrEmptyString(defaultTextOverride)) {
        defaultText = Office.Controls.Utils.htmlEncode(Office.Controls._peoplePickerTemplates._getDefaultText$i(allowMultiple));
    }
    else {
        defaultText = Office.Controls.Utils.htmlEncode(defaultTextOverride);
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
    var itemHtml = '<li tabindex=\"-1\" class=\"ms-PeoplePicker-result\" data-office-peoplepicker-value=\"' + Office.Controls.Utils.htmlEncode(principal.LoginName) + '\" title=\"' + titleText + '\">';
    itemHtml += '<div tabindex=\"-1\" class=\"ms-Persona ms-PersonaAdded\">';
    itemHtml += '<div tabindex=\"-1\" class=\"ms-Persona-details ms-Persona-detailsForDropdownAdded\">';
    itemHtml += '<a onclick=\"return false;\" href=\"#\" tabindex=\"-1\"><div tabindex=\"0\">';
    itemHtml += '<div class=\"ms-Persona-primaryText ms-Persona-primaryTextAdded\" >' + Office.Controls.Utils.htmlEncode(principal.DisplayName) + '</div>';
    if (!Office.Controls.Utils.isNullOrEmptyString(principal.JobTitle)) {
        itemHtml += '<div class=\"ms-Persona-secondaryText ms-Persona-secondaryTextAdded\" >' + Office.Controls.Utils.htmlEncode(principal.JobTitle) + '</div>';
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
    html += Office.Controls._peoplePickerTemplates._generateAutofillGroupTemplate$p(cachedEntries, 1, true);
    html += Office.Controls._peoplePickerTemplates._generateAutofillGroupTemplate$p(serverEntries, 0, false);
    html += '</div>';
    html += Office.Controls._peoplePickerTemplates.generateAutofillFooterTemplate(cachedEntries.length + serverEntries.length, maxCount);
    return html;
}
Office.Controls._peoplePickerTemplates._generateAutofillGroupTemplate$p = function(principals, source, isCached) {
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
    _re$p$0: null,
    sharePointHostUrl: null,
    appWebUrl: null,
    requestViaUrl: null,
    
    getRequestExecutor: function() {
        if (!this._re$p$0) {
            if (!Office.Controls.Utils.isNullOrEmptyString(Office.Controls.Runtime.context.appWebUrl)) {
                if (!Office.Controls.Utils.isNullOrEmptyString(Office.Controls.Runtime.context.requestViaUrl)) {
                    var options = new SP.RequestExecutorOptions();
                    options.viaUrl = Office.Controls.Runtime.context.requestViaUrl;
                    this._re$p$0 = new SP.RequestExecutor(Office.Controls.Runtime.context.sharePointHostUrl, options);
                }
                else {
                    this._re$p$0 = new SP.RequestExecutor(Office.Controls.Runtime.context.appWebUrl);
                }
            }
            else {
                Office.Controls.Utils.errorConsole('Missing authentication informations.');
            }
        }
        return this._re$p$0;
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
            var adapter = Access.ControlTelemetryAdapter.getStaticTelemetryAdapter(null, null);
            adapter.writeDiagnosticLog('EventListenerException', 'error', ex.message, null);
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
    for (var $$pai_8 = 1; $$pai_8 < arguments.length; ++$$pai_8) {
        args[$$pai_8 - 1] = arguments[$$pai_8];
    }
    var result = '';
    var i = 0;
    while (i < format.length) {
        var open = Office.Controls.Utils._findPlaceHolder$p(format, i, '{');
        if (open < 0) {
            result = result + format.substr(i);
            break;
        }
        else {
            var close = Office.Controls.Utils._findPlaceHolder$p(format, open, '}');
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
Office.Controls.Utils._findPlaceHolder$p = function(format, start, ch) {
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
//Office.Controls.PeoplePicker._searchPrincipalServerDataProvider.registerClass('Office.Controls.PeoplePicker._searchPrincipalServerDataProvider', null, Office.Controls.PeoplePicker.ISearchPrincipalDataProvider);
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
Office.Controls.PeoplePicker._autofillContainer._boolBodyHandlerAdded$p = false;
Office.Controls.PeoplePicker.ValidationError.multipleMatchName = 'MultipleMatch';
Office.Controls.PeoplePicker.ValidationError.multipleEntryName = 'MultipleEntry';
Office.Controls.PeoplePicker.ValidationError.noMatchName = 'NoMatch';
Office.Controls.PeoplePicker.ValidationError.serverProblemName = 'ServerProblem';
Office.Controls.PeoplePicker._mruCache._instance$p = null;
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
Office.Controls.PeoplePickerCustomerInsightStrings.loginName = 'LoginName';
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
