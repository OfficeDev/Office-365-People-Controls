(function () {
    "use strict";

    AadDataProvider = function () {
    };

    AadDataProvider.prototype = {
        maxResult: 50,
        lastErrorMessage: null,
        severHost: 'yihcaow1001:3000',
        getPrincipals: function (input, callback) {
            var xhr = new XMLHttpRequest(), self = this;
            xhr.open('GET', 'http://' + this.serverHost + '/api?keyword=' + encodeURIComponent(input), true);
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
        }
    };
})();



