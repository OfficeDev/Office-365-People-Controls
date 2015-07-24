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

    Office.Controls.Context = function (options) {
        if (typeof options !== 'object') {
            Office.Controls.Utils.errorConsole('Invalid parameters type');
            return;
        }
        var sharepointHost = options.HostUrl;
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
    Office.Controls.Runtime.initialize = function (options) {
        Office.Controls.Runtime.context = new Office.Controls.Context(options);
    };

    Office.Controls.Browser = function (browserType) {
        this.browserType = browserType;
        this.userAgent = navigator.userAgent.toLowerCase();
    } ;

    Office.Controls.Browser.prototype = {
        /**
         * Check the type of current browser
         * @return {Boolean}
         */
        isExpectedBrowser: function() {
            var isExpected = (this.userAgent.indexOf(this.browserType.toString()) !== -1);

            switch (this.browserType) {
                case Office.Controls.Browser.TypeEnum.IE:
                case Office.Controls.Browser.TypeEnum.Firefox:
                case Office.Controls.Browser.TypeEnum.Opera:
                    return isExpected;
                case Office.Controls.Browser.TypeEnum.Safari:
                    // The part of Chrome UserAgent value is AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36
                    return isExpected && (!this.isContainChromeStr());
                case Office.Controls.Browser.TypeEnum.Chrome:
                    // The part of Opera UserAgent value is Chrome/31.0.1650.63 Safari/537.36 OPR/18.0.1284.68
                    return isExpected && (!this.isContainOperaStr());
                default:
                    return false;
            }
        },

        isContainChromeStr: function() {
            return (this.userAgent.indexOf(Office.Controls.Browser.TypeEnum.Chrome.toString()) !== -1);
        },

        isContainOperaStr: function() {
            return (this.userAgent.indexOf(Office.Controls.Browser.TypeEnum.Opera.toString()) !== -1);
        }
    };

    // The browser type
    Office.Controls.Browser.TypeEnum = {
        // The article of search
        IE: "trident",
        // Table of contents, means the chapter of current search article
        Chrome: "chrome",
        // The images in current search article
        Firefox: "firefox",
        // The infobox tables in current search article
        Safari: "safari",
        // The reference of current search article
        Opera: "opr"
    };

    Office.Controls.Utils = function () { };
    // Construct browser in different browser type
    Office.Controls.Utils.isIE = function () {
        var browser = new Office.Controls.Browser(Office.Controls.Browser.TypeEnum.IE);
        return browser.isExpectedBrowser();
    };

    Office.Controls.Utils.isChrome = function () {
        var browser = new Office.Controls.Browser(Office.Controls.Browser.TypeEnum.Chrome);
        return browser.isExpectedBrowser();
    };

    Office.Controls.Utils.isSafari = function () {
        var browser = new Office.Controls.Browser(Office.Controls.Browser.TypeEnum.Safari);
        return browser.isExpectedBrowser();
    };

    Office.Controls.Utils.isFirefox = function () {
        var browser = new Office.Controls.Browser(Office.Controls.Browser.TypeEnum.Firefox);
        return browser.isExpectedBrowser();
    };

    Office.Controls.Utils.isOpera = function () {
        var browser = new Office.Controls.Browser(Office.Controls.Browser.TypeEnum.Opera);
        return browser.isExpectedBrowser();
    };

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

    Office.Controls.Utils.removeEventListener = function (element, eventName, handler) {
        var h = function (e) {
            try {
                return handler(e);
            } catch (ex) {
                throw ex;
            }
        };
        if (!Office.Controls.Utils.isNullOrUndefined(element.removeEventListener)) {
            element.removeEventListener(eventName, h, false);
        } else if (!Office.Controls.Utils.isNullOrUndefined(element.detachEvent )) {  // For IE
            element.detachEvent ('on' + eventName, h);
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
    Office.Controls.Utils.isFirefox = function () { return typeof InstallTrigger !== 'undefined' && navigator.userAgent.toLowerCase().indexOf('firefox') > -1; /* Firefox 1.0+ */ };
    Office.Controls.Utils.isIE10 = function () { return Function('/*@cc_on return document.documentMode===10@*/')(); } // jshint ignore:line
    Office.Controls.Utils.isFunction = function (functionToCheck) {
        var getType = {};
        return functionToCheck && getType.toString.call(functionToCheck) === '[object Function]';
    }
    Office.Controls.Utils.NOP = function () { };

    if (Office.Controls.Context.registerClass) { Office.Controls.Context.registerClass('Office.Controls.Context'); }
    if (Office.Controls.Runtime.registerClass) { Office.Controls.Runtime.registerClass('Office.Controls.Runtime'); }
    if (Office.Controls.Utils.registerClass) { Office.Controls.Utils.registerClass('Office.Controls.Utils'); }
    Office.Controls.Runtime.context = null;
    Office.Controls.Utils.oDataJSONAcceptString = 'application/json;odata=verbose';
    Office.Controls.Utils.clientTagHeaderName = 'X-ClientService-ClientTag';
})();