var aad = require("./aad");
var config = require("./config.js");
var superagent = require("superagent");
var baseuri = "https://graph.microsoft.com/beta"
var sectionsuri;
var refreshToken = "OAQABAAAAAACIv0qfZnbtS5u9YU9ubSAtS4i6jCF2iYXH86j--WbDkY3ft5Oq8Sm7-rbxOMSrluAdlTN9JNbxjekPgsudjMYi1LxAFF9WdhpxOu9PwHioa63-F4TzwjjHzj4XqQN0xEEyeCUfHI_hk1s4kJFnmrJyFhhjPTRM4Q0Lvi4vCJ7Id0-yTCK8hzhA064JU1JEP4eeaZw12jduFb2qI0tPfQGRhhrPxvThRx-LGVvLiTkWh2Jhm3qlM6DHypdHuHdZLw0kQ7qib66-q-hf_I5GkD4W2v8bwA1G6mEgX-HFYompdoEYk_cJD19I_g4qt6kfMh9FounT-oafkkx-ig1P4WRn9Yb4FA_KYTlhN-IOHpQbsVZZUYkrN5QvjgwPaTdEBe7YgMEwSq7IpMreS56Yflom4xnuqEhUL2cQlFWXepZlruFNqqpedQctNMrZ6A83Xx7mkdONzOcjoHFnYOpgem4p1ey9S4GmW17Flp_Optb5COyOiPWxzPu7lUjO7u-crV-0zaa4yUfFO-cnztuj65fmWU5frn7iw9SEwya7PqgRFMRW2_ATZXXnoa7IciYpS5Hls8RvEQM2yzdX582njuUmelMm6kfrqjq-DZJqlMckGCeakWj1xvEVLXMuDOQnK6GKJi9jLYADHJhcOcdinqUyX8WCkQaDxDQsEjqp3e1PuEaEHQGwiHAnK_gEwG-KKrSVOcnHgkY6w7rcYGB6zKH7n0p4w4kwYyFIokq9s9amCSbqaUqf8IHJarfirRd2ot0gAA";

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