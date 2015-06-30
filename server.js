'use strict';

// dependent modules
var fs = require('fs');
var path = require('path');
var http = require('http');
var https = require('https');
var crypto = require('crypto');
var express = require('express');
var through = require('through');
var cors = require('cors');
var engine = require('ejs-locals');
var cookie_parser = require('cookie-parser');
var express_session = require('express-session');
var AuthenticationContext = require('adal-node').AuthenticationContext;
var url = require('url');

http.globalAgent.maxSockets = Infinity;
https.globalAgent.maxSockets = Infinity;

// retrieve configuration values from azure env or config file.
var deploy_config = {};
if (fs.existsSync("config.json")) {
    deploy_config = JSON.parse(fs.readFileSync(path.resolve(__dirname, "config.json")));
}
var config_authorityHostUrl = process.env.AUTHORITYHOST || deploy_config.authorityHostUrl;
var config_site = process.env.SITEHOST || deploy_config.site;
var config_clientId = process.env.CLIENTID || deploy_config.clientId;
var config_clientSecret = process.env.CLIENTSECRET || deploy_config.clientSecret;

function sha256(str) {
    var sha256 = crypto.createHash("sha256");
    sha256.update(str, "utf8");
    return sha256.digest("base64");
}

// middleware for express
var app = express();
app.use(cors({
    origin: function(origin, callback) {
        var originIsWhitelisted = (typeof(process.env.ALLOWEDSITES) != 'undefined' && process.env.ALLOWEDSITES.indexOf(origin) !== -1)
            || (typeof(origin) != 'undefined' && origin.indexOf('http://localhost') !== -1);
        callback(null, originIsWhitelisted);
    },
    allowedHeaders: ['Content-Type', 'Authorization', 'Cookie'],
    credentials: true}));
app.use(cookie_parser('a deep secret'));
app.use(express_session({secret: '312569780QWE4TRY'}));
app.use(express.static(__dirname + '/static'));
app.engine('ejs', engine);
app.set('views', __dirname + '/views');
app.set('view engine', 'ejs');
app.use(express.static(__dirname + '/example'));

// constants
var authorityUrl = [config_authorityHostUrl, 'common'].join("/");
var redirectUri = ['http://', process.env.SITEHOST || config_site, '/accesstoken'].join("");
var ad_resource = '00000002-0000-0000-c000-000000000000';

var templateAuthzUrl = 'https://login.windows.net/common/oauth2/authorize?response_type=code&client_id=<client_id>&redirect_uri=<redirect_uri>&state=<state>&resource=<resource>';

function createAuthorizationUrl(state, resource, userType) {
    var authorizationUrl = templateAuthzUrl.replace('<client_id>', config_clientId);
    authorizationUrl = authorizationUrl.replace('<redirect_uri>', redirectUri);
    authorizationUrl = authorizationUrl.replace('<state>', state);
    authorizationUrl = authorizationUrl.replace('<resource>', resource);
    if (userType == "admin") {
        authorizationUrl += "&prompt=admin_consent"
    }
    console.log('/authcode authorizationUrl:' + authorizationUrl);
    return authorizationUrl;
}

// url mapping
app.get('/', function (req, res) {
    res.redirect('example/index.html');
});

app.get('/login', function (req, res) {
    res.redirect('/authcode');
});

app.get('/authcode', function (req, res) {
    try {
        crypto.randomBytes(48, function (ex, buf) {
            var auth_state = buf.toString('base64').replace(/\//g, '_').replace(/\+/g, '-');
            res.cookie('authstate', auth_state);
            res.cookie('redirect_uri', req.query.redirect_uri);
            console.log('/authcode redirect_uri:' + req.query.redirect_uri);
            console.log('/authcode authstate:' + auth_state);
            var authorizationUrl = createAuthorizationUrl(auth_state, ad_resource, req.query.userType);
            res.redirect(authorizationUrl);
        });
    } catch (e) {
        console.log('/authcode error:' + e);
    }
});

app.get('/accesstoken', function (req, res) {
    try {
        if (req.cookies.authstate !== req.query.state) {
            console.log('/accesstoken req.query.state:' + req.query.state);
            console.log('/accesstoken req.cookies.authstate:' + req.cookies.authstate);
            res.status(400).send('error: state does not match');
            return;
        }
        
        if (req.query.error) {
            if (typeof(req.cookies.redirect_uri) != 'undefined' && req.cookies.redirect_uri != 'undefined') {
                res.redirect(req.cookies.redirect_uri);
                return;
            }
        }

        var authenticationContext = new AuthenticationContext(authorityUrl);
        authenticationContext.acquireTokenWithAuthorizationCode(
            req.query.code,
            redirectUri,
            ad_resource,
            config_clientId,
            config_clientSecret,
            function (err, response) {
                var errorMessage = '';
                if (err) {
                    errorMessage = 'error: ' + err.message + '\n';
                    res.status(500).send(errorMessage);
                } else {
                    console.log('/accesstoken get by auth code');
                    //console.dir(response);
                    req.session.ad_tenantid = response.tenantId;
                    req.session.ad_accesstoken = [response.tokenType, response.accessToken].join(' ');

                    console.log('/accesstoken redirect_uri:' + req.cookies.redirect_uri);
                    console.log('/accesstoken ad_tenantid:' + response.tenantId);
                    console.log('/accesstoken ad_accesstoken:' + response.accessToken);
                    
                    if (typeof(req.cookies.redirect_uri) != 'undefined' && req.cookies.redirect_uri != 'undefined') {
                        res.redirect(req.cookies.redirect_uri + "?userId=" + response.userId);
                    } else {
                        //res.render('api');
                        res.status(200).send('Authenticated');
                    }
                }
        });
    } catch(e) {
      console.log('/accesstoken error:' + e);
    }
});

app.use('/users', function (req, res) {
    var header_buf = "";
    if (req.method === 'GET') {
        req.pipe(through(
        function (buf) {
            header_buf += buf.toString();
        }, function (buf) {
            var uploaded_header = header_buf.length === 0 ? req.headers : JSON.parse(header_buf);
            var tenantid = req.session.ad_tenantid;
            console.log('/users tenantid:' + tenantid);
            if (typeof(tenantid) == 'undefined') {
                res.status(401).send('Not authenticated.');
                return;
            }
            
            var keyword = req.query.keyword;
            var targeturl = 'https://graph.windows.net/' + tenantid + "/users?api-version=1.5";
            if (typeof(keyword) != undefined && keyword.length > 0) {
                targeturl += "&$filter=startswith(displayName," + encodeURIComponent("'" + keyword + "') or ")
                                + "startswith(userPrincipalName," + encodeURIComponent("'" + keyword + "')");
            }
            console.log('/users targeturl:'+targeturl);

            var header = {};
            header['Authorization'] = decodeURIComponent(req.session.ad_accesstoken);
            header['Accept'] = 'application/json;odata=nometadata;streaming=true;';

            https.get({
                host : url.parse(targeturl).host,
                path : url.parse(targeturl).path,
                headers : header
            }, function(response) {
              console.log("statusCode: ", response.statusCode);

              var res_body = "";
              response.on('data', function(d) {
                res_body += d;
              });
              response.on('end', function() {
                res.end(res_body);
              });

            }).on('error', function(e) {
                console.log("/users error: " + e);
                console.log('/users targeturl:' + targeturl);
                res.status(505).send();
            });
        }));
    }
});

app.use('/image', function (req, res) {
    var header_buf = "";
    if (req.method === 'GET') {
        req.pipe(through(
        function (buf) {
            header_buf += buf.toString();
        }, function (buf) {
            var uploaded_header = header_buf.length === 0 ? req.headers : JSON.parse(header_buf);
            var tenantid = req.session.ad_tenantid;
            console.log('/image tenantid:' + tenantid);
            if (typeof(tenantid) == 'undefined') {
                res.status(401).send('Not authenticated.');
                return;
            }
            var personId = req.query.personId;
            var targeturl = 'https://graph.windows.net/'+tenantid+'/users/'+personId+'/thumbnailPhoto?api-version=1.5';
            console.log('/image targeturl:'+targeturl);
            var header = {};
            header['Authorization'] = decodeURIComponent(req.session.ad_accesstoken);
            https.get({
                host : url.parse(targeturl).host,
                path : url.parse(targeturl).path,
                headers : header
            }, function(response) {
            console.log("statusCode: ", response.statusCode);
            if (response.statusCode != 200) {
                res.status(response.statusCode).send(null);
                return;
            }
            var buffers = [];
            response.on('data', function(d) {
            buffers.push(d);
            });
            response.on('end', function() {
                var image = Buffer.concat(buffers);
                res.end(image);
            });
            }).on('error', function(e) {
                console.log("/image error: " + e);
                res.status(505).send();
            });
        }));
    }
});

var port = process.env.PORT || 3000;

// start server
http.createServer(app).listen(port);
console.log('Server is listening on ' + port);
