# Repository Guidelines

## Project Structure & Module Organization
Egent is an Excel taskpane add-in. Core UI logic lives in `src/taskpane/components/` (see `App.tsx` for the Chat and Agent view switcher). Workbook helpers and Excel operations sit in `src/taskpane/utils/` and `src/taskpane/services/`, while `src/commands/` exposes the command callbacks referenced by the manifests in the repo root. Keep static images and markdown styles in `assets/`; production bundles land in `dist/` after builds.

## Build, Test, and Development Commands
Run `npm install` once per machine, then `npm run dev-server` to hot-reload the taskpane web app. Use `npm start` to sideload into Excel and surface the bottom dock described in the README; `npm stop` tears down debugging. Ship-ready assets come from `npm run build`, and `npm run validate` confirms `manifest.xml` before sharing. If you need to refresh Microsoft 365 auth, execute `npm run signin` or `npm run signout`.

## Coding Style & Naming Conventions
We code in TypeScript with React. Formatting is enforced by Office Add-in Prettier (`npm run prettier`) and lint rules (`npm run lint` or `npm run lint:fix`). Stick to two-space indentation, PascalCase component files, camelCase helpers, and kebab-case CSS. Follow the README examples when adding localized copy: default to English, but mirror existing Simplified Chinese strings when expanding UI text.

## Testing Guidelines
No automated harness exists yet, so rely on workbook smoke tests. Validate both modes: Chat must respect read-only behavior, while Agent may mutate ranges via Excel function calling (for example, re-selecting a stored range). Document the sheets, ranges, and prompts you used in each PR. When you add automated coverage, colocate `*.test.tsx` beside the feature and mock Office.js calls.

## Commit & Pull Request Guidelines
Recent history favors concise, feature-focused summaries, often in Chinese (for example, `优化 markdown 样式`). Keep subjects under 72 characters and use the imperative mood. Pull requests should link any tracked issues, outline validation commands, attach UI screenshots of the taskpane tabs, and note changes to manifests or OpenAI configuration expectations.

## Agent & Chat Mode Notes
Chat mode behaves like GitHub Copilot for Excel: it reads the active selection and replies without altering cells. Agent mode can insert text, recolor ranges, or re-apply selections through the helper functions in `src/taskpane/taskpane.ts`. Ensure new features clearly state which mode they require and guard destructive actions with confirmation prompts.
