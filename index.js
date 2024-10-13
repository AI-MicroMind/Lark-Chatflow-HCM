const express = require("express");
const dotenv = require("dotenv");
const fetch = require("node-fetch");

dotenv.config();

const app = express();
const port = process.env.PORT || 3000;

const FLOWISE_API_URL = process.env.FLOWISE_API_URL || "";

app.use(express.json());

const conversationHistories = {};

function logger(...params) {
  console.log(`[Teams Integration]`, ...params);
}

async function query(data, sessionId) {
  try {
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
    const result = await response.json();

    history.push(result.text);
    conversationHistories[sessionId] = history;

    return result;
  } catch (error) {
    logger("Error querying Flowise API:", error);
    throw error;
  }
}

async function handleMessage(text, sessionId) {
  logger("Received message:", text);
  
  if (text.toLowerCase().startsWith("/clear")) {
    delete conversationHistories[sessionId];
    return "✅ Conversation history cleared.";
  }

  const response = await query({ question: text }, sessionId);
  return response.text;
}

app.post("/teams-webhook", async (req, res) => {
  logger("Received webhook request:", JSON.stringify(req.body, null, 2));
  
  let text, sessionId;

  // Handle different possible message formats
  if (req.body.type === "message") {
    text = req.body.text;
    sessionId = req.body.conversation.id;
  } else if (req.body.text && req.body.sessionId) {
    text = req.body.text;
    sessionId = req.body.sessionId;
  } else {
    logger("Unrecognized message format");
    return res.status(400).json({ error: "Invalid input format" });
  }

  if (!text || !sessionId) {
    logger("Missing text or sessionId");
    return res.status(400).json({ error: "Invalid input. 'text' and 'sessionId' are required." });
  }

  try {
    const answer = await handleMessage(text, sessionId);
    logger("Sending response:", answer);
    res.json({ 
      type: "message",
      text: answer
    });
  } catch (error) {
    logger("Error handling Teams webhook:", error);
    res.status(500).json({ error: "Internal Server Error. Please try again later." });
  }
});

app.get("/health", (req, res) => {
  res.json({ status: "UP", message: "Teams webhook is running" });
});

app.listen(port, () => {
  logger(`Teams webhook server running on port ${port}`);
});
