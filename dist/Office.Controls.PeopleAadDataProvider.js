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

    Office.Controls.PeopleAadDataProvider = function (authContext) {
        this.authContext = authContext;
    }

    Office.Controls.PeopleAadDataProvider.prototype = {
        maxResult: 50,
        authContext: null,
        lastErrorMessage: null,
        severHost: 'localhost:3000', // local debug
        acquireToken: '00000002-0000-0000-c000-000000000000', // AAD resource key
        apiVersion: 'api-version=1.5', 

        getPrincipals: function (keyword, callback) {
            var xhr = new XMLHttpRequest(), self = this;
            xhr.open('GET', 'http://' + this.serverHost + '/api?keyword=' + encodeURIComponent(keyword), true);
            xhr.setRequestHeader('Content-Type', 'application/json');
            xhr.withCredentials = true;
            xhr.onabort = xhr.onerror = xhr.ontimeout = function () {
                self.lastErrorMessage = 'Search error. Try login first.';
                callback('Error', null);
            };
            xhr.onload = function () {
                if (xhr.status === 401) {
                    self.lastErrorMessage = 'Unauthorized. You need login first.';
                    callback('Unauthorized', null);
                    return;
                }
                if (xhr.status !== 200) {
                    self.lastErrorMessage = 'Unknown error. Status code: ' + xhr.statusCode;
                    callback('Unknown error', null);
                }
                var result = JSON.parse(xhr.responseText), people = [];
                if (result["odata.error"] !== undefined) {
                    self.lastErrorMessage = 'Server error: ' + result["odata.error"].message.value;
                    callback(result["odata.error"], null);
                    return;
                }
                result.value.forEach(
                    function (e) {
                        var person = {};
                        person.displayName = e.displayName;
                        person.department = e.department;
                        person.jobTitle = e.jobTitle;
                        person.mail = e.mail;
                        person.workPhone = e.telephoneNumber;
                        person.mobile = e.mobile;
                        person.office = e.physicalDeliveryOfficeName;
                        person.sipAddress = e.userPrincipalName;
                        person.alias = e.mailNickname;
                        person.personId = e.objectId;
                        people.push(person);
                    });
                if (people.length > self.maxResult) {
                    people = people.slice(0, self.maxResult);
                }
                callback(null, people);
            };
            xhr.send('');
        },

        /**
         * CORS
         * @param  {[type]}
         * @param  {Function}
         * @return {[type]}
         */
        getUsersAsync: function (keyword, callback) {

            var self = this;
            self.authContext.acquireToken(self.acquireToken, function (error, token) {

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
                xhr.open('GET', 'https://graph.windows.net/' + tenant + '/users?' + apiVersion + "&$filter=startswith(displayName," +
                    encodeURIComponent("'" + keyword + "') or ") + "startswith(userPrincipalName," + encodeURIComponent("'" + keyword + "')"), true);
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
                            person.displayName = e.displayName;
                            person.department = e.department;
                            person.jobTitle = e.jobTitle;
                            person.mail = e.mail;
                            person.workPhone = e.telephoneNumber;
                            person.mobile = e.mobile;
                            person.office = e.physicalDeliveryOfficeName;
                            person.sipAddress = e.userPrincipalName;
                            person.alias = e.mailNickname;
                            person.personId = e.objectId;
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

        getUserImageAsync: function (personId, callback) {

            var self = this;
            self.authContext.acquireToken(self.acquireToken, function (error, token) {

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
                xhr.open('GET', 'https://graph.windows.net/' + tenant + '/users/' + personId + '/thumbnailPhoto?' + apiVersion);
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
})();