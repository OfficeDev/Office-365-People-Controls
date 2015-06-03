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
        // console.error(message);
    };

    Office.Controls.Utils.getObjectFromFullyQualifiedName = function (objectName) {
        var currentObject = window.self;
        return Office.Controls.Utils.getObjectFromJSONObjectName(currentObject, objectName);
    };

    // Parse the json object to get the corresponding value
    Office.Controls.Utils.getObjectFromJSONObjectName = function (jsonObject, objectName) {
        var currentObject = jsonObject;
        if (Office.Controls.Utils.isNullOrUndefined(currentObject)) {
                return null;
        } 

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

    if (Office.Controls.Context.registerClass) { Office.Controls.Context.registerClass('Office.Controls.Context'); }
    if (Office.Controls.Runtime.registerClass) { Office.Controls.Runtime.registerClass('Office.Controls.Runtime'); }
    if (Office.Controls.Utils.registerClass) { Office.Controls.Utils.registerClass('Office.Controls.Utils'); }
    Office.Controls.Runtime.context = null;
    Office.Controls.Utils.oDataJSONAcceptString = 'application/json;odata=verbose';
    Office.Controls.Utils.clientTagHeaderName = 'X-ClientService-ClientTag';
})();