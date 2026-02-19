import {
  stripMentionsText,
  TokenCredentials,
  ConversationReference,
  toActivityParams,
} from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { LocalStorage } from "@microsoft/teams.common";
import config from "./config";
import { ManagedIdentityCredential } from "@azure/identity";

// Create storage for conversation history
const storage = new LocalStorage();

// Store conversation references for proactive messaging
const conversationReferences = new Map<string, ConversationReference>();

const createTokenFactory = () => {
  return async (scope: string | string[], tenantId?: string): Promise<string> => {
    const managedIdentityCredential = new ManagedIdentityCredential({
      clientId: process.env.CLIENT_ID,
    });
    const scopes = Array.isArray(scope) ? scope : [scope];
    const tokenResponse = await managedIdentityCredential.getToken(scopes, {
      tenantId: tenantId,
    });

    return tokenResponse.token;
  };
};

// Configure authentication using TokenCredentials
const tokenCredentials: TokenCredentials = {
  clientId: process.env.CLIENT_ID || "",
  token: createTokenFactory(),
};

const credentialOptions =
  config.MicrosoftAppType === "UserAssignedMsi" ? { ...tokenCredentials } : undefined;

// Create the app with storage
const app = new App({
  ...credentialOptions,
  storage,
});

// Capture conversation reference on bot install
app.on("install.add" as any, async (context: any) => {
  conversationReferences.set(context.activity.conversation.id, context.ref);
  console.log(`[message] conversationId on install: ${context.activity.conversation.id}`);
  console.log(`[message] ref:`, JSON.stringify(context.ref, null, 2));
});

// Interface for conversation state
interface ConversationState {
  count: number;
}

const getConversationState = (conversationId: string): ConversationState => {
  let state = storage.get(conversationId);
  if (!state) {
    state = { count: 0 };
    storage.set(conversationId, state);
  }
  return state;
};

app.on("message", async (context) => {
  const activity = context.activity;

  // Store/update conversation reference on every message
  console.log(`[message] conversationId: ${activity.conversation.id}`);
  console.log(`[message] ref:`, JSON.stringify(context.ref, null, 2));
  conversationReferences.set(activity.conversation.id, context.ref);

  const text: string = stripMentionsText(activity);

  if (text === "/reset") {
    storage.delete(activity.conversation.id);
    await context.send("Ok I've deleted the current conversation state.");
    return;
  }

  if (text === "/count") {
    const state = getConversationState(activity.conversation.id);
    await context.send(`The count is ${state.count}`);
    return;
  }

  if (text === "/diag") {
    await context.send(JSON.stringify(activity));
    return;
  }

  if (text === "/state") {
    const state = getConversationState(activity.conversation.id);
    await context.send(JSON.stringify(state));
    return;
  }

  if (text === "/runtime") {
    const runtime = {
      nodeversion: process.version,
      sdkversion: "2.0.0", // Microsoft Teams SDK
    };
    await context.send(JSON.stringify(runtime));
    return;
  }

  // Default echo behavior
  const state = getConversationState(activity.conversation.id);
  state.count++;
  await context.send(`[${state.count}] you said: ${text}`);
});

// --- Proactive messaging API ---

// POST /api/notify - Send a proactive message to a conversation
// Body: {
//   "conversationId": "<id>",
//   "message": "<text>",
//   "replyToId": "<optional message id for threading>",
//   "mentions": [{ "id": "<user AAD id>", "name": "<display name>" }]
// }
// To mention someone, include <at>Name</at> in the message text and add a
// corresponding entry in the mentions array.
app.http.post("/api/notify", async (req: any, res: any) => {
  const { conversationId, message, replyToId, mentions } = req.body || {};

  if (!conversationId || !message) {
    res.status(400).json({ error: "conversationId and message are required" });
    return;
  }

  const ref = conversationReferences.get(conversationId);
  if (!ref) {
    res.status(404).json({
      error: "Conversation not found. The bot must receive a message from this conversation first.",
      knownConversations: Array.from(conversationReferences.keys()),
    });
    return;
  }

  try {
    const activityParams = toActivityParams(message);

    // Build mention entities if provided
    if (mentions && Array.isArray(mentions) && mentions.length > 0) {
      activityParams.entities = mentions.map((m: { id: string; name: string }) => ({
        type: "mention",
        text: `<at>${m.name}</at>`,
        mentioned: {
          id: m.id,
          name: m.name,
        },
      }));
    }

    // If replyToId is provided, send as a threaded reply in a channel
    // Teams requires the conversation ID to include the message ID for threading
    if (replyToId) {
      const threadRef: ConversationReference = {
        ...ref,
        conversation: {
          ...ref.conversation,
          id: `${ref.conversation.id};messageid=${replyToId}`,
        },
      };
      const result = await app.http.send(activityParams, threadRef);
      res.json({ status: "sent", activityId: result.id });
    } else {
      const result = await app.http.send(activityParams, ref);
      res.json({ status: "sent", activityId: result.id });
    }
  } catch (err: any) {
    res.status(500).json({ error: err.message || "Failed to send message" });
  }
});

// GET /api/conversations - List known conversation IDs
app.http.get("/api/conversations" as any, async (_req: any, res: any) => {
  const conversations = Array.from(conversationReferences.entries()).map(
    ([id, ref]) => ({
      conversationId: id,
      conversationType: ref.conversation?.conversationType,
      serviceUrl: ref.serviceUrl,
    })
  );
  res.json({ conversations });
});

export default app;
