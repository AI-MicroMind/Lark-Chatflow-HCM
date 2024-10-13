import express from "express";
import dotenv from "dotenv";
import fetch from "node-fetch";

dotenv.config();

const app = express();
const port = process.env.PORT || 3000;

const FLOWISE_API_URL = process.env.FLOWISE_API_URL || "";

app.use(express.json());

const conversationHistories = {};

function logger(...params) {
  console.error(`[Teams Integration]`, ...params);
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
    console.error("Error querying Flowise API:", error);
    throw error;
  }
}

async function handleMessage(text, sessionId) {
  logger("Received message:", text);
  
  if (text.startsWith("/clear")) {
    delete conversationHistories[sessionId];
    return "✅ Conversation history cleared.";
  }

  const response = await query({ question: text }, sessionId);
  return response.text;
}

app.post("/teams-webhook", async (req, res) => {
  const { text, sessionId } = req.body;

  if (!text || !sessionId) {
    return res.status(400).json({ error: "Invalid input. 'text' and 'sessionId' are required." });
  }

  try {
    const answer = await handleMessage(text, sessionId);
    res.json({ message: answer });
  } catch (error) {
    logger("Error handling Teams webhook:", error);
    res.status(500).json({ error: "Internal Server Error. Please try again later." });
  }
});

app.get("/health", (req, res) => {
  res.json({ status: "UP", message: "Teams webhook is running" });
});

app.listen(port, () => {
  console.log(`Teams webhook server running on port ${port}`);
});
