var restify = require('restify');
var builder = require('botbuilder');
var api = require('./api.js');
var _ = require("underscore")
//var auth = require('./oauth.js');
var request = require('request');
var text;



// Create bot and add dialogs
var bot = new builder.TextBot({ appId: 'OneNoteTextBot', appSecret: 'f698da3c4c2f4839ac7c911bb969c7fb' });
bot.add('/', new builder.CommandDialog()
    .matches('^yes', builder.DialogAction.beginDialog('/Yes'))
    .matches('^no', builder.DialogAction.beginDialog('/No') )
    .onDefault([
        function (session, results){
            text = session.message.text;
            session.send("You want to send this to your Quicknote? yes/no ?")
           // console.log(results.response);
        }
        
    ]));

bot.listenStdin();

bot.add('/Yes', [
    function (session) {
        Go();
        builder.Prompts.text(session, 'Your QuickNote Page has been updated');
    },
    function (session, results) {
        session.userData.name = results.response;
        session.endDialog();
    }
]);

bot.add('/No', [
    function (session) {
        builder.Prompts.text(session, 'Your QuickNote Page has not been updated');
    },
    function (session, results) {
        session.userData.name = results.response;
        session.endDialog();
    }
]);

// Setup Restify Server
var server = restify.createServer();
server.post('/api/messages', bot.verifyBotFramework(), bot.listen());

// Serve a static web page
server.get(/.*/, restify.serveStatic({
	'directory': '.',
	'default': 'index.html'
}));


server.listen(process.env.port || 3978, function () {
    console.log('%s listening to %s', server.name, server.url); 
});


//Talk to OneNote

function Go() {
api.getSections(function(err, result){
    if(err)
        console.log(err);
   else{
       length = result.value.length - 1;
       api.getPages(result.value[length].id, function(err, result){
           if(err)
            console.log(err);
           else{
               console.log(result.value[length].id);
               api.updatePages(result.value[length].id, text, function(err, result){
                   if(err)
                     console.log(err);
                   else{
                       console.log("SUCCESS")
                   }
               } )
           }
       });
    }
});
}
