const DEFAULT_API_BASE = "https://text.pollinations.ai/openai";
const DEFAULT_CHAT_PATH = "/v1/chat/completions";
const DEFAULT_MODEL = "openai";

function json(statusCode, payload) {
  return {
    statusCode,
    headers: {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "no-store",
    },
    body: JSON.stringify(payload),
  };
}

function toMessages(body) {
  if (Array.isArray(body.messages) && body.messages.length > 0) {
    return body.messages;
  }
  if (typeof body.prompt === "string" && body.prompt.trim() !== "") {
    return [{ role: "user", content: body.prompt.trim() }];
  }
  return null;
}

function extractReply(data) {
  const choice = data?.choices?.[0];
  if (!choice) return "";
  if (typeof choice.text === "string") return choice.text;
  if (typeof choice?.message?.content === "string") return choice.message.content;
  return "";
}

export async function handler(event) {
  if (event.httpMethod === "OPTIONS") {
    return {
      statusCode: 204,
      headers: {
        "access-control-allow-methods": "POST, OPTIONS",
        "access-control-allow-headers": "content-type",
      },
      body: "",
    };
  }

  if (event.httpMethod !== "POST") {
    return json(405, { error: "Method not allowed. Use POST." });
  }

  const apiKey = process.env.POLLINATIONS_API_KEY;

  let body = {};
  try {
    body = event.body ? JSON.parse(event.body) : {};
  } catch {
    return json(400, { error: "Invalid JSON body." });
  }

  const messages = toMessages(body);
  if (!messages) {
    return json(400, {
      error: "Request must include `prompt` or `messages`.",
    });
  }

  const apiBase =
    (process.env.POLLINATIONS_API_BASE_URL || DEFAULT_API_BASE).replace(/\/+$/u, "");
  const chatPath = process.env.POLLINATIONS_CHAT_PATH || DEFAULT_CHAT_PATH;
  const url = `${apiBase}${chatPath.startsWith("/") ? chatPath : `/${chatPath}`}`;

  const requestedModel = String(
    body.model || process.env.POLLINATIONS_MODEL || DEFAULT_MODEL,
  );

  const payload = {
    model: requestedModel,
    messages,
    stream: false,
  };

  if (typeof body.temperature === "number") payload.temperature = body.temperature;
  if (typeof body.max_tokens === "number") payload.max_tokens = body.max_tokens;
  if (typeof body.maxTokens === "number") payload.max_tokens = body.maxTokens;
  if (typeof body.top_p === "number") payload.top_p = body.top_p;
  if (typeof body.presence_penalty === "number") payload.presence_penalty = body.presence_penalty;
  if (typeof body.frequency_penalty === "number") payload.frequency_penalty = body.frequency_penalty;

  let upstream;
  try {
    const headers = {
      "content-type": "application/json",
    };
    if (apiKey) {
      headers.authorization = `Bearer ${apiKey}`;
    }

    upstream = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
    });
  } catch (error) {
    return json(502, {
      error: "Failed to reach Pollinations API.",
      detail: error instanceof Error ? error.message : String(error),
    });
  }

  const raw = await upstream.text();
  let data;
  try {
    data = raw ? JSON.parse(raw) : {};
  } catch {
    data = { raw };
  }

  if (!upstream.ok) {
    return json(upstream.status, {
      error: "Pollinations API returned an error.",
      upstream: data,
      meta: {
        endpoint: url,
        requestedModel,
        responseModel:
          typeof data === "object" && data !== null && "model" in data
            ? data.model ?? null
            : null,
      },
    });
  }

  return json(200, {
    reply: extractReply(data),
    data,
    meta: {
      endpoint: url,
      requestedModel,
      responseModel:
        typeof data === "object" && data !== null && "model" in data
          ? data.model ?? null
          : null,
    },
  });
}
