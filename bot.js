const { ActionTypes, ActivityTypes, CardFactory } = require('botbuilder');
const path = require('path');
const axios = require('axios');
const fs = require('fs');
const fetch = require('node-fetch');
var BizCard = require('./bizCardInput.json')

class AttachmentsBot {
    /**
     * Every conversation turn for our AttachmentsBot will call this method.
     * There are no dialogs used, since it's "single turn" processing, meaning a single
     * request and response, with no stateful conversation.
     * @param turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */

    async onTurn(turnContext) {
        if (turnContext.activity.type === ActivityTypes.Message) {
            // Determine how the bot should process the message by checking for attachments.
            if (turnContext.activity.attachments && turnContext.activity.attachments.length > 0) {
                // The user sent an attachment and the bot should handle the incoming attachment.
                await this.handleIncomingAttachment(turnContext);
            } else {
                // Since no attachment was received, send an attachment to the user.
                await this.handleOutgoingAttachment(turnContext);
            }
            // Send a HeroCard with potential options for the user to select.
            await this.displayOptions(turnContext);
        } else if (turnContext.activity.type === ActivityTypes.ConversationUpdate &&
            turnContext.activity.recipient.id !== turnContext.activity.membersAdded[0].id) {
            // If the Activity is a ConversationUpdate, send a greeting message to the user.
            await turnContext.sendActivity('Welcome to the Business Card Handling Bot! Send me an business card and I will save it.');
            // Send a HeroCard with potential options for the user to select.
            await this.displayOptions(turnContext);
        } else if (turnContext.activity.type !== ActivityTypes.ConversationUpdate) {
            // Respond to all other Activity types.
            await turnContext.sendActivity(`[${ turnContext.activity.type }]-type activity detected.`);
        }
    }

    /**
     * Saves incoming attachments to disk by calling `this.downloadAttachmentAndWrite()` and
     * responds to the user with information about the saved attachment or an error.
     * @param {Object} turnContext
     */
    async handleIncomingAttachment(turnContext) {
        
        // Prepare Promises to download each attachment and then execute each Promise.
        const promises = turnContext.activity.attachments.map(this.downloadAttachmentAndWrite);
        const successfulSaves = await Promise.all(promises);

        // Replies back to the user with information about where the attachment is stored on the bot's server,
        // and what the name of the saved file is.
        async function replyForReceivedAttachments(localAttachmentData) {
            if (localAttachmentData) {
                // Because the TurnContext was bound to this function, the bot can call
                // `TurnContext.sendActivity` via `this.sendActivity`;
                let bizCardDetails = await getBizCard(localAttachmentData.base64);
                console.log(bizCardDetails);
                let bizCardRes = buildCard(bizCardDetails, localAttachmentData.localPath);
                await this.sendActivity({
                    text: 'Here is recognized details:',
                    attachments: [CardFactory.adaptiveCard(bizCardRes)]
                });
            } else {
                await this.sendActivity('Attachment was not successfully saved to disk.');
            }
        }

        // Prepare Promises to reply to the user with information about saved attachments.
        // The current TurnContext is bound so `replyForReceivedAttachments` can also send replies.
        const replyPromises = successfulSaves.map(replyForReceivedAttachments.bind(turnContext));
        await Promise.all(replyPromises);
    }

    /**
     * Downloads attachment to the disk.
     * @param {Object} attachment
     */
    async downloadAttachmentAndWrite(attachment) {
        // Retrieve the attachment via the attachment's contentUrl.
        
        const url = attachment.contentUrl;

        // Local file path for the bot to save the attachment.
        const localFileName = path.join(__dirname, "images\\" + attachment.name);

        try {
            // arraybuffer is necessary for images
            const response = await axios.get(url, { responseType: 'arraybuffer' });
            // If user uploads JSON file, this prevents it from being written as "{"type":"Buffer","data":[123,13,10,32,32,34,108..."
            if (response.headers['content-type'] === 'application/json') {
                response.data = JSON.parse(response.data, (key, value) => {
                    return value && value.type === 'Buffer' ?
                      Buffer.from(value.data) :
                      value;
                    });
            }
            attachment.base64 = Buffer.from(response.data).toString('base64');
            fs.writeFile(localFileName, response.data, (fsError) => {
                if (fsError) {
                    throw fsError;
                }
            });
        } catch (error) {
            console.error(error);
            return undefined;
        }

        // If no error was thrown while writing to disk, return the attachment's name
        // and localFilePath for the response back to the user.
        return {
            fileName: attachment.name,
            localPath: localFileName,
            base64: attachment.base64
        };
    }

    /**
     * Responds to user with either an attachment or a default message indicating
     * an unexpected input was received.
     * @param {Object} turnContext
     */
    async handleOutgoingAttachment(turnContext) {
        const reply = { type: ActivityTypes.Message };

        // Look at the user input, and figure out what type of attachment to send.
        // If the input matches one of the available choices, populate reply with
        // the available attachments.
        // If the choice does not match with a valid choice, inform the user of
        // possible options.
        const firstChar = turnContext.activity.text[0];
        if (firstChar === '1') {
            await this.recognizeCard(turnContext);
        } else if (firstChar === '2') {
            await this.searchCard(turnContext);
        } else {
            // The user did not enter input that this bot was built to handle.
            reply.text = 'Your input was not recognized, please try again.';
        }
        await turnContext.sendActivity(reply);
    }

    /**
     * Sends a HeroCard with choices of attachments.
     * @param {Object} turnContext
     */
    async displayOptions(turnContext) {
        const reply = { type: ActivityTypes.Message };

        // Note that some channels require different values to be used in order to get buttons to display text.
        // In this code the emulator is accounted for with the 'title' parameter, but in other channels you may
        // need to provide a value for other parameters like 'text' or 'displayText'.
        const buttons = [
            { type: ActionTypes.ImBack, title: '1. Upload Business Card', value: '1' },
            { type: ActionTypes.ImBack, title: '2. Search Contacts', value: '2' }
        ];

        const card = CardFactory.heroCard('', undefined,
            buttons, { text: 'Please select one of the following choices.' });

        reply.attachments = [card];

        await turnContext.sendActivity(reply);
    }

    /**
     * Returns an inline attachment.
     */
    async searchCard(turnContext) {
        await turnContext.sendActivity("Search Card Function Triggered");
    }

    /**
     * Returns an attachment to be sent to the user from a HTTPS URL.
     */
    async recognizeCard(turnContext) {
        await turnContext.sendActivity("Recognized Card Function Triggered");
    }

}

exports.AttachmentsBot = AttachmentsBot;

getBizCard = async function(img) {
    req = {'image' : img}
    let res = await fetch('https://warrenjsfunc.azurewebsites.net/api/BusinessCardOCR?code=ZZ4e7PdyaaIav/0/7RdhKMZIrLf6SbkJklU27XS0fMa5I6jDx8Qsxg==', {
            method: 'post',
            body:    JSON.stringify(req),
        })
    return(res.json());
};

buildCard = (cardDetails, path) => {
    res = BizCard;
    for (i in BizCard.body[0].columns[0].items) {
        if (BizCard.body[0].columns[0].items[i].type == "Input.Text") {
            switch (BizCard.body[0].columns[0].items[i].id) {
                case "myName":
                    res.body[0].columns[0].items[i].value = cardDetails.name;
                    break;
                case "myEmail":
                    res.body[0].columns[0].items[i].value = cardDetails.email;
                    break;
                case "myTel":
                    res.body[0].columns[0].items[i].value = "12345678";
                    break;
                case "myOrg":
                    res.body[0].columns[0].items[i].value = cardDetails.organization;
                    break;
                case "myWeb":
                    res.body[0].columns[0].items[i].value = cardDetails.website;
                    break;
                default:
                    break;
            }
        }
        if (BizCard.body[0].columns[0].items[i].type == "Image") {
            res.body[0].columns[0].items[i].url = path;
        }
    }
    return res;
}
