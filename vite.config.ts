import type { IncomingMessage, ServerResponse } from "node:http";
import path from "node:path";
import { defineConfig, loadEnv, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

const DEFAULT_API_BASE = "https://text.pollinations.ai/openai";
const DEFAULT_CHAT_PATH = "/v1/chat/completions";
const DEFAULT_MODEL = "openai";

function sendJson(res: ServerResponse, statusCode: number, payload: unknown) {
  res.statusCode = statusCode;
  res.setHeader("content-type", "application/json; charset=utf-8");
  res.setHeader("cache-control", "no-store");
  res.end(JSON.stringify(payload));
}

function readBody(req: IncomingMessage) {
  return new Promise<string>((resolve, reject) => {
    let body = "";
    req.on("data", (chunk) => {
      body += chunk;
    });
    req.on("end", () => resolve(body));
    req.on("error", reject);
  });
}

function toMessages(body: Record<string, unknown>) {
  if (Array.isArray(body.messages) && body.messages.length > 0) {
    return body.messages;
  }
  if (typeof body.prompt === "string" && body.prompt.trim() !== "") {
    return [{ role: "user", content: body.prompt.trim() }];
  }
  return null;
}

function extractReply(data: any) {
  const choice = data?.choices?.[0];
  if (!choice) return "";
  if (typeof choice.text === "string") return choice.text;
  if (typeof choice?.message?.content === "string") return choice.message.content;
  return "";
}

function pollinationsDevProxy(env: Record<string, string>): Plugin {
  return {
    name: "pollinations-dev-proxy",
    configureServer(server) {
      server.middlewares.use(async (req, res, next) => {
        const pathname = req.url?.split("?")[0] ?? "";
        if (pathname !== "/api/pollinations-chat") {
          return next();
        }

        if (req.method === "OPTIONS") {
          res.statusCode = 204;
          res.setHeader("access-control-allow-methods", "POST, OPTIONS");
          res.setHeader("access-control-allow-headers", "content-type");
          return res.end();
        }

        if (req.method !== "POST") {
          return sendJson(res, 405, { error: "Method not allowed. Use POST." });
        }

        const apiKey = env.POLLINATIONS_API_KEY || process.env.POLLINATIONS_API_KEY;

        let body: Record<string, unknown> = {};
        try {
          const rawBody = await readBody(req);
          body = rawBody ? (JSON.parse(rawBody) as Record<string, unknown>) : {};
        } catch {
          return sendJson(res, 400, { error: "Invalid JSON body." });
        }

        const messages = toMessages(body);
        if (!messages) {
          return sendJson(res, 400, {
            error: "Request must include `prompt` or `messages`.",
          });
        }

        const apiBase = (
          env.POLLINATIONS_API_BASE_URL ||
          process.env.POLLINATIONS_API_BASE_URL ||
          DEFAULT_API_BASE
        ).replace(/\/+$/u, "");
        const chatPath =
          env.POLLINATIONS_CHAT_PATH ||
          process.env.POLLINATIONS_CHAT_PATH ||
          DEFAULT_CHAT_PATH;
        const url = `${apiBase}${chatPath.startsWith("/") ? chatPath : `/${chatPath}`}`;

        const requestedModel = String(
          body.model ||
            env.POLLINATIONS_MODEL ||
            process.env.POLLINATIONS_MODEL ||
            DEFAULT_MODEL,
        );

        const payload: Record<string, unknown> = {
          model: requestedModel,
          messages,
          stream: false,
        };

        if (typeof body.temperature === "number") payload.temperature = body.temperature;
        if (typeof body.max_tokens === "number") payload.max_tokens = body.max_tokens;
        if (typeof body.maxTokens === "number") payload.max_tokens = body.maxTokens;
        if (typeof body.top_p === "number") payload.top_p = body.top_p;
        if (typeof body.presence_penalty === "number")
          payload.presence_penalty = body.presence_penalty;
        if (typeof body.frequency_penalty === "number")
          payload.frequency_penalty = body.frequency_penalty;

        let upstream;
        try {
          const headers: Record<string, string> = {
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
          return sendJson(res, 502, {
            error: "Failed to reach Pollinations API.",
            detail: error instanceof Error ? error.message : String(error),
          });
        }

        const raw = await upstream.text();
        let data: unknown;
        try {
          data = raw ? JSON.parse(raw) : {};
        } catch {
          data = { raw };
        }

        if (!upstream.ok) {
          return sendJson(res, upstream.status, {
            error: "Pollinations API returned an error.",
            upstream: data,
            meta: {
              endpoint: url,
              requestedModel,
              responseModel:
                typeof data === "object" && data !== null && "model" in data
                  ? (data as { model?: string }).model ?? null
                  : null,
            },
          });
        }

        return sendJson(res, 200, {
          reply: extractReply(data),
          data,
          meta: {
            endpoint: url,
            requestedModel,
            responseModel:
              typeof data === "object" && data !== null && "model" in data
                ? (data as { model?: string }).model ?? null
                : null,
          },
        });
      });
    },
  };
}

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), "");

  return {
    plugins: [react(), tailwindcss(), pollinationsDevProxy(env)],
    resolve: {
      dedupe: ["react", "react-dom", "react/jsx-runtime"],
      alias: {
        react: path.resolve(__dirname, "node_modules/react"),
        "react-dom": path.resolve(__dirname, "node_modules/react-dom"),
      },
    },
    optimizeDeps: {
      include: ["react", "react-dom", "react/jsx-runtime", "react-dom/client"],
    },
    server: {
      hmr: {
        protocol: "ws",
        host: "localhost",
      },
    },
  };
});
