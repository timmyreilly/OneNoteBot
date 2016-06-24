var aad = require("./aad");
var config = require("./config.js");
var superagent = require("superagent");
var baseuri = "https://graph.microsoft.com/beta"
var env = require('./env.js');
var sectionsuri;
var refresh = process.env.ACCESS_TOKEN;
var accessToken;


function getToken(cb) {
     if (!accessToken) {
        aad.refreshToken(refreshToken, config, function(err, x){
          if(err)
            cb(err);
          else
            cb(null, x.access_token);
            
        });
     } else {
       cb(null, accessToken);
     }
}

function getSections(cb){
  getToken(function (err, token) {
    if (err)
      cb(err);
    else {
      //console.log(token);
      superagent
        .get(baseuri + '/me/notes/sections')
        .set("Authorization", token)
        .set("Accept", "application/json")
        .end((err, result) => {
          if(err)
            cb(err)
          else
            cb(null, JSON.parse(result.text))
        });
    }
  });
  
}

function getPages(id, cb){
  getToken(function (err, token) {
    if (err)
      cb(err);
    else {
      sectionsuri = baseuri + '/me/notes/sections/' + id + '/pages';
      superagent
        .get(baseuri + '/me/notes/sections/' + id + '/pages')
        .set("Authorization", token)
        .set("Accept", "application/json")
        .end((err, result) => {
          if(err)
            cb(err)
          else
            cb(null, JSON.parse(result.text))
        });
    }
  });
  
}

function updatePages(id, text, cb){
  getToken(function (err, token) {
    if (err)
      cb(err);
    else {
      console.log(baseuri + "/me/notes/pages/" + id + '/content');
      superagent
        .patch(baseuri + "/me/notes/pages/" + id + '/content')
        .send([{
          'action':'append',
          'target': 'body',
          'content': text,
        }])
        .set("Authorization", token)
        .set("Content-Type", "application/json")
        .end((err, result) => {
          if(err)
            cb(err)
          else
            cb(null, result)
        });
    }
  });
  
}

exports.getSections = getSections;
exports.getPages = getPages;
exports.updatePages = updatePages;