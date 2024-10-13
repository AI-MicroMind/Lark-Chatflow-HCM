const express = require('express');
const dotenv = require('dotenv');
const { BotFrameworkAdapter, ActivityTypes } = require('botbuilder');
const fetch = require('node-fetch');


dotenv.config();

const app = express();
const port = process.env.PORT || 3000;

const FLOWISE_API_URL = process.env.FLOWISE_API_URL || "";
const MICROSOFT_APP_ID = process.env.MICROSOFT_APP_ID;
const MICROSOFT_APP_PASSWORD = process.env.MICROSOFT_APP_PASSWORD;

// Create bot adapter
const adapter = new BotFrameworkAdapter({
    appId: MICROSOFT_APP_ID,
    appPassword: MICROSOFT_APP_PASSWORD
});

const conversationHistories = {};

function logger(...params) {
    console.log(`[Teams Integration]`, ...params);
}

async function query(data, sessionId) {
    try {
        logger(`Querying Flowise API with: ${JSON.stringify(data)}`);
        const history = conversationHistories[sessionId] || [];
        history.push(data.question);
        conversationHistories[sessionId] = history;

        const response = await fetch(FLOWISE_API_URL, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify({ 
                question: history.join(" "),
                overrideConfig: {
                    sessionId: sessionId
                }
            }),
        });

        if (!response.ok) {
            throw new Error(`Flowise API responded with status: ${response.status}`);
        }

        const result = await response.json();
        logger(`Flowise API response: ${JSON.stringify(result)}`);

        if (!result.text) {
            throw new Error('Unexpected response format from Flowise API');
        }

        history.push(result.text);
        conversationHistories[sessionId] = history;

        return result.text;
    } catch (error) {
        logger("Error querying Flowise API:", error);
        throw error;
    }
}

async function handleMessage(text, sessionId) {
    logger("Handling message:", text);
    
    if (text.toLowerCase().startsWith("/clear")) {
        delete conversationHistories[sessionId];
        return "✅ Conversation history cleared.";
    }

    return await query({ question: text }, sessionId);
}

// Handle incoming activities
app.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    if (context.activity.type === ActivityTypes.Message) {
      if (context.activity.attachments && context.activity.attachments.length > 0) {
        // Handle file attachments
        for (let attachment of context.activity.attachments) {
          if (attachment.contentType.startsWith('image/') || attachment.contentType.startsWith('video/')) {
            // Process the file
            await handleFileAttachment(context, attachment);
          }
        }
      } else {
        // Handle text messages as before
        const text = context.activity.text;
        const sessionId = context.activity.conversation.id;
        const reply = await handleMessage(text, sessionId);
        await context.sendActivity(reply);
      }
    }
  });
});

async function handleFileAttachment(context, attachment) {
  // Download the file
  const fileDownload = await axios.get(attachment.contentUrl, { responseType: 'arraybuffer' });
  
  // Process the file (this is where you'd integrate with your Flowise API or other services)
  // For now, we'll just acknowledge receipt
  await context.sendActivity(`Received ${attachment.contentType} file: ${attachment.name}`);
}

app.get("/health", (req, res) => {
    res.json({ status: "UP", message: "Teams webhook is running" });
});

app.listen(port, () => {
    logger(`Microsoft Teams bot server running on port ${port}`);
});
