/**
 * @module routes/api/sense-bot/microsoft
 * @author yianni.ververis@qlik.com
 * @description
 * Handle all of the https://{domain}/api/sense-bot/microsoft/ routes
 * skype: https://join.skype.com/bot/bc69f77c-331a-4775-808b-4346866f082f
 * skype-svc: https://join.skype.com/bot/514357d9-b843-449f-8a1d-a1dfbf065978?add
*/

const express = require('express');
const site = require('../../../../models/sense-bot');
const builder = require('botbuilder');
const router = express.Router();
const config = require('../../../../config.json');
const shared = require('./shared');
const teams = require('botbuilder-teams');
const inMemoryStorage = new builder.MemoryBotStorage(); //Refer to https://blog.botframework.com/2017/12/19/bot-state-service-will-soon-retired-march-31st-2018/ for more  details 
let lang = 'en';
let qvf = config.qvf.salesforce;
shared.engine = new site.Enigma(qvf);

// Create chat connector for communicating with the Bot Framework Service
// let connector = new builder.ChatConnector({
// 	appId: process.env.MICROSOFT_APP_ID,
// 	appPassword: process.env.MICROSOFT_APP_PASSWORD
// });
let connector = new teams.TeamsChatConnector({
	// It is a bad idea to store secrets in config files. We try to read the settings from
	// the config file (/config/default.json) OR then environment variables.
	// See node config module (https://www.npmjs.com/package/config) on how to create config files for your Node.js environment.
	// appId: config.get("bot.appId"),
	// appPassword: config.get("bot.appPassword")
	appId: process.env.MICROSOFT_APP_ID,
	appPassword : process.env.MICROSOFT_APP_PASSWORD
});


// let bot = new builder.UniversalBot(connector);
// let bot = new builder.UniversalBot(connector, {storage: new builder.MemoryBotStorage()}, [ function (session) {
let bot = new builder.UniversalBot(connector, [async function (session) {
	try {
		lang = 'en';//session.preferredLocale();
		prompt = session.localizer.gettext(session.preferredLocale(), "text_prompt");
		//Store the user for Sending Messages later
		// let db = await new site.Microsoft();
		// let result = await db.userListing({
		// 	userUid: session.message.user.id,
		// 	limit: 1
		// });
		// if (!result.length) {
		// 	let channelId = 1;
		// 	if (session.message.address.channelId === 'msteams') {
		// 		channelId = 2;
		// 	} else if (session.message.address.channelId === 'cortana') {
		// 		channelId = 3;
		// 	} else if (session.message.address.channelId === 'webchat') {
		// 		channelId = 5;
		// 	}
		// 	await db.userInsert({
		// 		userUid: session.message.user.id,
		// 		username: session.message.user.name,
		// 		channelId: channelId,
		// 		userData: JSON.stringify(session.message.address).trim()
		// 	});
		// }
		// // Open Dialogs based on the text the user types
		if (
			session.message.text === "helpdesk" ||
			session.message.text === "cio" ||
			session.message.text === "salesforce" ||
			session.message.text === "help"
		) {
			session.send(`Hi! You are connected to ${session.message.text}. What do you want to do?`);
			session.send(session.message.address);
			session.beginDialog(session.message.text);
		} else {
			session.send(config.text[lang].no_command, session.message.text);
			session.beginDialog('help');
		}
		site.logger.info(`loaded`, { route: `api/sense-bot/microsoft::UniversalBot` });
	}
	catch (error) {
		site.logger.info(`error: ${error}`, { route: `api/sense-bot/microsoft::UniversalBot()` });
	}
}]).set('storage', inMemoryStorage);

bot.on('contactRelationUpdate', function (message) {
	try {
		if (message.action === 'add') {
			var reply = new builder.Message()
				.address(message.address)
				.text(config.text[lang].welcome, message.user ? message.user.name : 'there');
			bot.send(reply);
		}
		site.logger.info(`loaded`, { route: `api/sense-bot/microsoft::contactRelationUpdate()` });
	}
	catch (error) {
		site.logger.info(`error: ${error}`, { route: `api/sense-bot/microsoft::contactRelationUpdate()` });
	}
});

// Send welcome when conversation with bot is started, by initiating the root dialog
bot.on('conversationUpdate', function (message) {
	try {
		if (message.membersAdded && message.membersAdded.length > 0) {
			var reply = new builder.Message()
				.address(message.address)
				.text(config.text[lang].welcome, message.user ? message.user.name : 'there');
			bot.send(reply);
		} else if (message.membersRemoved) {
			// See if bot was removed
			var botId = message.address.bot.id;
			for (var i = 0; i < message.membersRemoved.length; i++) {
				if (message.membersRemoved[i].id === botId) {
					var reply = new builder.Message()
						.address(message.address)
						.text(config.text[lang].exit.text);
					bot.send(reply);
					break;
				}
			}
		}
		site.logger.info(`loaded`, { route: `api/sense-bot/microsoft::conversationUpdate()` });
	}
	catch (error) {
		site.logger.info(`error: ${error}`, { route: `api/sense-bot/microsoft::conversationUpdate()` });
	}
});

// Exit from all dialogs
bot.dialog('exit', [function (session) {
	try {
		if (shared.engine) { shared.engine.disconnect(); }
		session.sendTyping();
		session.endDialog(config.text[lang].exit.text);
		session.beginDialog('help');
		site.logger.info(`loaded`, { route: `api/sense-bot/microsoft::exit()` });
	}
	catch (error) {
		site.logger.info(`error: ${error}`, { route: `api/sense-bot/microsoft::exit()` });
	}
}])
	.triggerAction({ matches: /^exit$/i });

// This is used in the field manipulation tasks 
let appName = require('./help')(bot, builder) || "salesforce";
// Other app specific bot dialogs called
require('./locale')(bot, builder);
require('./salesforce')(bot, builder);

// clear selections 
bot.dialog('clearSelection', [function (session) {
	try {
		if (!shared.engine) { 'You need to be connected to an app to clear selections!' };
		shared.engine.clear();
		session.sendTyping();
		site.logger.info(`loaded`, { route: `api/sense-bot/microsoft::clearedSelections()` });
		session.endDialog("Selections cleared");
		session.beginDialog(appName);   // Need to make it dynamic based on user app selected
	}
	catch (error) {
		site.logger.info(`error: ${error}`, { route: `api/sense-bot/microsoft::clearedSelections()` });
		session.beginDialog(appName);  
	}
}])
.triggerAction({ matches: /^clear selections$/i });

// get fields
bot.dialog('getFields', [ async function (session) {
	try {
		session.sendTyping();
		if (!shared.engine) { 'You need to be connected to an app to getfields!' };
		site.logger.info(`loaded`, { route: `api/sense-bot/microsoft::getFields()` });	
		let fieldValues = await shared.getFields();   //Array
		let fieldValueActionCard = [];
		fieldValues.forEach( async (fieldValue, index) => {
			fieldValueActionCard.push(		
				builder.CardAction.postBack(session, fieldValue, `data:${fieldValue}`)
			);
		});
		let msg = await new builder.Message(session);
		msg.attachmentLayout(builder.AttachmentLayout.list);
		msg.attachments([
			new builder.HeroCard(session)
				.title("Dimensions")
				.text(config.text.dimension_action)
				.buttons(fieldValueActionCard)
		]);		
		session.send(msg);
		//session.beginDialog(appName);   // Need to make it dynamic based on user app selected
	}
	catch (error) {
		site.logger.info(`error: ${error}`, { route: `api/sense-bot/microsoft::getFields()` });
		session.beginDialog(appName);  
	}
}])
	.triggerAction({ matches: /^get fields$/i });

// select field data
bot.dialog('selectValue', [ async function (session) {
	try {
		session.sendTyping();
		if (!shared.engine) { 'You need to be connected to an app to clear selections!' };
		site.logger.info(`loaded`, { route: `api/sense-bot/microsoft::selectFieldValue()` });		
		let [ fieldName, fieldValue ]  = session.message.text.split(',');
		if(!fieldValue) {
			 session.send("Please select value to select\n");
			 session.endDialog( await shared.getFieldData(fieldName.replace(/select/i, '').trim()) );
			 session.beginDialog(appName); 
		}else {  
		let msg = await shared.select( fieldName.replace(/select/i, " ").trim(), fieldValue.trim());
		session.endDialog(`${fieldValue.trim()} is selected in ${fieldName.replace(/select/i, " ").trim()}`); 
		session.beginDialog(appName);   // Need to make it dynamic based on user app selected
		}
	}
	catch (error) {
		site.logger.info(`error: ${error}`, { route: `api/sense-bot/microsoft::selectFieldValue()` });
		session.beginDialog(appName);  
	}
}])
	.triggerAction( { matches: /select [a-z0-9\s]*,[a-z0-9\s]*/i } );	// yet to design regualar expression 

// get field data
bot.dialog('getFieldValue', [ async function (session) {
	try {
		if (!shared.engine) { 'You need to be connected to an app to clear selections!' };
		site.logger.info(`loaded`, { route: `api/sense-bot/microsoft::getFieldValues()` });
		session.sendTyping();
		let fieldName = session.message.text.split(':')[1].trim();
		session.endDialog(await shared.getFieldData(fieldName));
		session.beginDialog(appName);   // Need to make it dynamic based on user app selected
	}
	catch (error) {
		site.logger.info(`error: ${error}`, { route: `api/sense-bot/microsoft::getFieldValues()` });
		session.beginDialog(appName);  
	}
}])
	.triggerAction({ matches: /^data:./i });	





//----------------------------------------------------------------------------------------------------------------

bot.dialog('SendO365Card', function (session) {
    // multiple choice examples
    var actionCard1 = new teams.O365ConnectorCardActionCard(session)
        .id("card-1")
        .name("Multiple Choice")
        .inputs([
        new teams.O365ConnectorCardMultichoiceInput(session)
            .id("list-1")
            .title("Pick multiple options")
            .isMultiSelect(true)
            .isRequired(true)
            .style('expanded')
            .choices([
            new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Choice 1").value("1"),
            new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Choice 2").value("2"),
            new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Choice 3").value("3")
        ]),
        new teams.O365ConnectorCardMultichoiceInput(session)
            .id("list-2")
            .title("Pick multiple options")
            .isMultiSelect(true)
            .isRequired(true)
            .style('compact')
            .choices([
            new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Choice 4").value("4"),
            new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Choice 5").value("5"),
            new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Choice 6").value("6")
        ]),
        new teams.O365ConnectorCardMultichoiceInput(session)
            .id("list-3")
            .title("Pick an options")
            .isMultiSelect(false)
            .style('expanded')
            .choices([
            new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Choice a").value("a"),
            new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Choice b").value("b"),
            new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Choice c").value("c")
        ]),
        new teams.O365ConnectorCardMultichoiceInput(session)
            .id("list-4")
            .title("Pick an options")
            .isMultiSelect(false)
            .style('compact')
            .choices([
            new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Choice x").value("x"),
            new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Choice y").value("y"),
            new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Choice z").value("z")
        ])
    ])
        .actions([
        new teams.O365ConnectorCardHttpPOST(session)
            .id("card-1-btn-1")
            .name("Send")
            .body(JSON.stringify({
            list1: '{{list-1.value}}',
            list2: '{{list-2.value}}',
            list3: '{{list-3.value}}',
            list4: '{{list-4.value}}'
        }))
    ]);
    // text input examples
    var actionCard2 = new teams.O365ConnectorCardActionCard(session)
        .id("card-2")
        .name("Text Input")
        .inputs([
        new teams.O365ConnectorCardTextInput(session)
            .id("text-1")
            .title("multiline, no maxLength")
            .isMultiline(true),
        new teams.O365ConnectorCardTextInput(session)
            .id("text-2")
            .title("single line, no maxLength")
            .isMultiline(false),
        new teams.O365ConnectorCardTextInput(session)
            .id("text-3")
            .title("multiline, max len = 10, isRequired")
            .isMultiline(true)
            .isRequired(true)
            .maxLength(10),
        new teams.O365ConnectorCardTextInput(session)
            .id("text-4")
            .title("single line, max len = 10, isRequired")
            .isMultiline(false)
            .isRequired(true)
            .maxLength(10)
    ])
        .actions([
        new teams.O365ConnectorCardHttpPOST(session)
            .id("card-2-btn-1")
            .name("Send")
            .body(JSON.stringify({
            text1: '{{text-1.value}}',
            text2: '{{text-2.value}}',
            text3: '{{text-3.value}}',
            text4: '{{text-4.value}}'
        }))
    ]);
    // date / time input examples
    var actionCard3 = new teams.O365ConnectorCardActionCard(session)
        .id("card-3")
        .name("Date Input")
        .inputs([
        new teams.O365ConnectorCardDateInput(session)
            .id("date-1")
            .title("date with time")
            .includeTime(true)
            .isRequired(true),
        new teams.O365ConnectorCardDateInput(session)
            .id("date-2")
            .title("date only")
            .includeTime(false)
            .isRequired(false)
    ])
        .actions([
        new teams.O365ConnectorCardHttpPOST(session)
            .id("card-3-btn-1")
            .name("Send")
            .body(JSON.stringify({
            date1: '{{date-1.value}}',
            date2: '{{date-2.value}}'
        }))
    ]);
    var section = new teams.O365ConnectorCardSection(session)
        .markdown(true)
        .title("**section title**")
        .text("section text")
        .activityTitle("activity title")
        .activitySubtitle("activity sbtitle")
        .activityImage("http://connectorsdemo.azurewebsites.net/images/MSC12_Oscar_002.jpg")
        .activityText("activity text")
        .facts([
        new teams.O365ConnectorCardFact(session).name("Fact name 1").value("Fact value 1"),
        new teams.O365ConnectorCardFact(session).name("Fact name 2").value("Fact value 2"),
    ])
        .images([
        new teams.O365ConnectorCardImage(session).title("image 1").image("http://connectorsdemo.azurewebsites.net/images/MicrosoftSurface_024_Cafe_OH-06315_VS_R1c.jpg"),
        new teams.O365ConnectorCardImage(session).title("image 2").image("http://connectorsdemo.azurewebsites.net/images/WIN12_Scene_01.jpg"),
        new teams.O365ConnectorCardImage(session).title("image 3").image("http://connectorsdemo.azurewebsites.net/images/WIN12_Anthony_02.jpg")
	]);
	
    var card = new teams.O365ConnectorCard(session)
        .summary("O365 card summary")
        .themeColor("#E67A9E")
        .title("card title")
        .text("card text")
        .sections([section])
        .potentialAction([
        actionCard1,
        actionCard2,
        actionCard3,
        new teams.O365ConnectorCardViewAction(session)
            .name('View Action')
            .target('http://microsoft.com'),
        new teams.O365ConnectorCardOpenUri(session)
            .id('open-uri')
            .name('Open Uri')["default"]('http://microsoft.com')
            .iOS('http://microsoft.com')
            .android('http://microsoft.com')
            .windowsPhone('http://microsoft.com')
    ]);
    var msg = new teams.TeamsMessage(session)
 //       .summary("A sample O365 actionable card")
        .attachments([card]);
	session.send(msg);
	
	session.send(`Hi ${session.message.user.name.split(' ')[0]}`);
	session.endDialog();
}).triggerAction({ matches: /^test/i });
// // example for o365 connector actionable card
// var o365CardActionHandler = function (event, query, callback) {
//     var userName = event.address.user.name;
//     var body = JSON.parse(query.body);
//     var msg = new builder.Message()
//         .address(event.address)
//         .summary("Thanks for your input!")
//         .textFormat("xml")
//         .text("<h2>Thanks, " + userName + "!</h2><br/><h3>Your input action ID:</h3><br/><pre>" + query.actionId + "</pre><br/><h3>Your input body:</h3><br/><pre>" + JSON.stringify(body, null, 2) + "</pre>");
//     connector.send([msg.toMessage()], function (err, address) {
//     });


//-----------------------------------------------------------------------------------------------------------------








// SKYPE SENSE BOT POST MESSAGES
router.post('/', connector.listen());

// POST MESSAGE TO ALL USERS
router.post('/adhoc/', async (req, res) => {
	try {
		let db = await new site.Microsoft();
		let result = await db.userListing({ all: true });
		if (result.length == 1) {
			let value = result[0].user_data;
			value = value.replace("\\", "");
			value = JSON.parse(value);
			var msg = new builder.Message().address(value);
			msg.text(req.body.message);
			msg.textLocale('en-US');
			bot.send(msg);
		} else if (result.length > 1) {
			for (let value of result) {
				var msg = new builder.Message().address(JSON.parse(value.user_data));
				msg.text(req.body.message);
				msg.textLocale('en-US');
				bot.send(msg);
			}
		}
		res.send({
			success: true,
			data: `Message "${req.body.message}" Send!`
		});
		site.logger.info(`adhoc`, { route: `${req.originalUrl}` });
		site.logger.info(`adhoc-message`, `${req.body.message}`);
	}
	catch (error) {
		site.logger.info(`adhoc-error`, { route: `${JSON.stringify(error)}` });
		res.send({
			success: false,
			data: `Error: "${JSON.stringify(error)}"`
		});
	}
});

module.exports = router;