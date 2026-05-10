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
  timeoutMs?: number;
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

const DEFAULT_ENDPOINT = "/api/pollinations-chat";

const POLLINATIONS_ENDPOINT =
  import.meta.env.VITE_POLLINATIONS_ENDPOINT || DEFAULT_ENDPOINT;

export async function requestPollinations(
  input: PollinationsRequest,
): Promise<PollinationsResponse> {
  const { timeoutMs = 0, maxTokens, ...payload } = input;
  const controller =
    timeoutMs > 0 && typeof AbortController !== "undefined"
      ? new AbortController()
      : undefined;
  let timeoutId: ReturnType<typeof setTimeout> | undefined;

  if (controller) {
    timeoutId = setTimeout(() => controller.abort(), timeoutMs);
  }

  let response: Response;
  try {
    response = await fetch(POLLINATIONS_ENDPOINT, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        ...payload,
        max_tokens: maxTokens,
      }),
      signal: controller?.signal,
    });
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") {
      throw new Error("AI request timed out.");
    }
    throw error;
  } finally {
    if (timeoutId !== undefined) {
      clearTimeout(timeoutId);
    }
  }

  const raw = await response.text();
  let json: {
    reply?: string;
    data?: unknown;
    meta?: {
      endpoint?: string;
      requestedModel?: string | null;
      responseModel?: string | null;
    };
    error?: string;
    detail?: string;
  } = {};

  try {
    json = raw ? (JSON.parse(raw) as typeof json) : {};
  } catch {
    json = {};
  }

  if (!response.ok) {
    const msg = json.error || `Pollinations request failed (${response.status}).`;
    const detail =
      typeof json.detail === "string" && json.detail.trim() !== ""
        ? ` ${json.detail}`
        : "";
    throw new Error(`${msg}${detail}`);
  }

  return {
    reply: typeof json.reply === "string" ? json.reply : "",
    data: json.data,
    meta: json.meta,
  };
}
