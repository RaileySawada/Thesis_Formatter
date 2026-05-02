# React + TypeScript + Vite

This template provides a minimal setup to get React working in Vite with HMR and some ESLint rules.

Currently, two official plugins are available:

- [@vitejs/plugin-react](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react) uses [Oxc](https://oxc.rs)
- [@vitejs/plugin-react-swc](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react-swc) uses [SWC](https://swc.rs/)

## React Compiler

The React Compiler is not enabled on this template because of its impact on dev & build performances. To add it, see [this documentation](https://react.dev/learn/react-compiler/installation).

## Expanding the ESLint configuration

If you are developing a production application, we recommend updating the configuration to enable type-aware lint rules:

```js
export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      // Other configs...

      // Remove tseslint.configs.recommended and replace with this
      tseslint.configs.recommendedTypeChecked,
      // Alternatively, use this for stricter rules
      tseslint.configs.strictTypeChecked,
      // Optionally, add this for stylistic rules
      tseslint.configs.stylisticTypeChecked,

      // Other configs...
    ],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.node.json', './tsconfig.app.json'],
        tsconfigRootDir: import.meta.dirname,
      },
      // other options...
    },
  },
])
```

## Pollinations AI on Netlify

This project includes a secure Netlify Function proxy:

- Function route: `/.netlify/functions/pollinations-chat`
- Function file: `netlify/functions/pollinations-chat.mjs`
- Client helper: `src/lib/pollinationsClient.ts`
- Local dev route: `/api/pollinations-chat` (served by Vite middleware)

### Netlify environment variables

Set these in Netlify (`Site settings -> Environment variables`):

- `POLLINATIONS_API_KEY` (required)
- `POLLINATIONS_API_BASE_URL` (optional, default `https://text.pollinations.ai/openai`)
- `POLLINATIONS_CHAT_PATH` (optional, default `/v1/chat/completions`)
- `POLLINATIONS_MODEL` (optional, default `openai`)

### Local development

1. Create `.env` in the project root (or copy from `.env.example`).
2. Set `POLLINATIONS_API_KEY=your_key_here`.
3. Optional: set `VITE_ENABLE_AI_ASSIST=true` to enable an AI health-check before formatting.
4. Run `npm run dev`.

The frontend helper auto-selects:

- `/api/pollinations-chat` in local dev
- `/.netlify/functions/pollinations-chat` in production

### Security

Never put this API key in client-side `VITE_*` variables.  
Use server-side function environment variables only.

You can also install [eslint-plugin-react-x](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-x) and [eslint-plugin-react-dom](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-dom) for React-specific lint rules:

```js
// eslint.config.js
import reactX from 'eslint-plugin-react-x'
import reactDom from 'eslint-plugin-react-dom'

export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      // Other configs...
      // Enable lint rules for React
      reactX.configs['recommended-typescript'],
      // Enable lint rules for React DOM
      reactDom.configs.recommended,
    ],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.node.json', './tsconfig.app.json'],
        tsconfigRootDir: import.meta.dirname,
      },
      // other options...
    },
  },
])
```
