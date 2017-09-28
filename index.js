var express = require('express'); 
var fs = require('fs');
var app = express();
app.get('/', function(req, res) {
    res.sendFile('home.html', {root: __dirname })
}
