type Role = "system" | "user" | "assistant";

export interface PollinationsMessage {
  role: Role;
  content: string;
}

export interface PollinationsRequest {
  prompt?: string;
  messages?: PollinationsMessage[];
  model?: string;
  temperature?: number;
  maxTokens?: number;
}

export interface PollinationsResponse {
  reply: string;
  data: unknown;
  meta?: {
    endpoint?: string;
    requestedModel?: string | null;
    responseModel?: string | null;
  };
}

const DEFAULT_ENDPOINT = import.meta.env.DEV
  ? "/api/pollinations-chat"
  : "/.netlify/functions/pollinations-chat";

const POLLINATIONS_ENDPOINT =
  import.meta.env.VITE_POLLINATIONS_ENDPOINT || DEFAULT_ENDPOINT;

export async function requestPollinations(
  input: PollinationsRequest,
): Promise<PollinationsResponse> {
  const response = await fetch(POLLINATIONS_ENDPOINT, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      ...input,
      max_tokens: input.maxTokens,
    }),
  });

  const json = (await response.json()) as {
    reply?: string;
    data?: unknown;
    meta?: {
      endpoint?: string;
      requestedModel?: string | null;
      responseModel?: string | null;
    };
    error?: string;
  };

  if (!response.ok) {
    throw new Error(json.error || "Pollinations request failed.");
  }

  return {
    reply: typeof json.reply === "string" ? json.reply : "",
    data: json.data,
    meta: json.meta,
  };
}
