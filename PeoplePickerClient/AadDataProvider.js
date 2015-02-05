(function () {
    AadDataProvider = function() {
    }

    AadDataProvider.prototype = {
        lastErrorMessage: null,
        severHost: 'yihcaow1001:3000',
        getPrincipals: function (input, callback) {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', 'http://' + this.serverHost + '/api?keyword=' + encodeURIComponent(input), true);
            xhr.setRequestHeader('Content-Type', 'application/json');
            xhr.withCredentials = true;
            var self = this;
            xhr.onabort = xhr.onerror = xhr.ontimeout = function () {
                self.lastErrorMessage = 'Search error. Try login first.';
                callback('Error', null);
            };
            xhr.onload = function () {
                if (xhr.statusCode == 401) {
                    self.lastErrorMessage = 'Unauthorized. You need login first.';
                    callback('Unauthorized', null);
                    return;
                }
                else if (xhr.statusCode != 200) {
                    self.lastErrorMessage = 'Unknown error.';
                    callback('Unknown error', null);
                }
                var result = JSON.parse(xhr.responseText);
                if ("odata.error" in result) {
                    self.lastErrorMessage = 'Server error: ' + result["odata.error"]["message"]["value"];
                    callback(result["odata.error"], null);
                    return;
                }
                var people = new Array();
                result["value"].forEach(
                    function (e) {
                        var person = {};
                        person.DisplayName = e.displayName;
                        person.Description = e.department;
                        person.PersonId = e.objectId;
                        people.push(person);
                    });
                callback(null, people);
            };
            xhr.send('');
        }
    }
})();



