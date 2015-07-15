(function () {

    AadDataProvider = function () {
    }

    AadDataProvider.prototype = {
        maxResult: 50,
        lastErrorMessage: null,
        severHost: 'localhost:3000',
        getImageAsync: function (personId, callback) {
            var xhr = new XMLHttpRequest(), self = this;
            xhr.open('GET', 'http://' + this.serverHost + '/image?personId=' + encodeURIComponent(personId), true);
            xhr.setRequestHeader('Content-Type', 'application/json');
            xhr.withCredentials = true;
            xhr.responseType = "blob";
            
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
                    return;
            }
            var reader = new FileReader();
            reader.addEventListener("loadend", function() {
                callback(null, reader.result);
            });
            reader.readAsDataURL(xhr.response);
            };
            xhr.send('');
        },
        searchPeopleAsync: function (keyword, callback) {
            var xhr = new XMLHttpRequest(), self = this;
            xhr.open('GET', 'http://' + this.serverHost + '/users?keyword=' + encodeURIComponent(keyword), true);
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
                    return;
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
                        person.description = person.department = e.department;
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
        }
    };
})();



