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
const backupHistories = {};

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
    
    if (text.toLowerCase() === "/clear" || text.toLowerCase() === "/new") {
        if (conversationHistories[sessionId]) {
            backupHistories[sessionId] = backupHistories[sessionId] || [];
            backupHistories[sessionId].push([...conversationHistories[sessionId]]);
        }
        delete conversationHistories[sessionId];
        return "✅ Conversation history cleared and backed up. You can start a new session now.";
    }

    if (text.toLowerCase() === "/history") {
        const backedUpChats = backupHistories[sessionId];
        if (!backedUpChats || backedUpChats.length === 0) {
            return "No backed-up chat history available.";
        }
        return "Here are your backed-up chats:\n\n" + 
               backedUpChats.map((chat, index) => 
                 `Chat ${index + 1}:\n${chat.join("\n")}\n`
               ).join("\n");
    }

    const response = await query({ question: text }, sessionId);
    
    if (response.toLowerCase().includes("task complete") || response.toLowerCase().includes("anything else i can help you with")) {
        if (conversationHistories[sessionId]) {
            backupHistories[sessionId] = backupHistories[sessionId] || [];
            backupHistories[sessionId].push([...conversationHistories[sessionId]]);
        }
        delete conversationHistories[sessionId];
        return response + "\n\nThe session has been cleared and backed up. You can start a new inquiry.";
    }

    return response;
}

// Handle incoming activities
app.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        if (context.activity.type === ActivityTypes.Message) {
            const text = context.activity.text;
            const sessionId = context.activity.conversation.id;

            try {
                const reply = await handleMessage(text, sessionId);
                await context.sendActivity(reply);
            } catch (error) {
                logger("Error processing message:", error);
                await context.sendActivity("I'm sorry, I encountered an error while processing your request.");
            }
        }
    });
});

app.get("/health", (req, res) => {
    res.json({ status: "UP", message: "Teams webhook is running" });
});

app.listen(port, () => {
    logger(`Microsoft Teams bot server running on port ${port}`);
});
