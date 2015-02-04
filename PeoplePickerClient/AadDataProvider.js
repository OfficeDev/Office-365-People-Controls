
var AadDataProvider = {
    serverHost: 'yihcaow1001:3000',
    getPrincipals: function (input, callback) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'http://' + this.serverHost + '/api?keyword=' + encodeURIComponent(input), true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.withCredentials = true;
        xhr.onabort = xhr.onerror = xhr.ontimeout = function () {
            callback('Error', null);
        };
        xhr.onload = function () {
            if (xhr.statusCode == 401) {
                callback('Unauthorized', null);
                return;
            }
            var result = JSON.parse(xhr.responseText);
            console.log(' result:' + xhr.responseText);
            var persons = new Array();
            result["value"].forEach(
                function (e) {
                    var person = {};
                    person.DisplayName = e.displayName;
                    person.Description = e.department;
                    person.PersonId = e.objectId;
                    persons.push(person);
                });
            callback(null, persons);
        };
        xhr.send('');
    }
};



