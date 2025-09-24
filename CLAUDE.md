# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Egent is a Microsoft Office Excel add-in built as a task pane application using React, TypeScript, and the Office.js API. The project aims to provide AI agent functionality for Excel users, offering both chat-based interactions and automated Excel data manipulation through OpenAI Function Calling.

## Development Commands

### Build and Development
- `npm run build` - Production build using webpack
- `npm run build:dev` - Development build
- `npm run dev-server` - Start webpack dev server on port 3000
- `npm run watch` - Development build with file watching
- `npm start` - Start Office add-in debugging with manifest.xml
- `npm stop` - Stop Office add-in debugging

### Code Quality
- `npm run lint` - Run Office add-in linting checks
- `npm run lint:fix` - Auto-fix linting issues
- `npm run prettier` - Format code using Office add-in prettier config
- `npm run validate` - Validate the manifest.xml file

### Office 365 Authentication
- `npm run signin` - Sign in to M365 account for development
- `npm run signout` - Sign out of M365 account

## Architecture

### Entry Points
- **TaskPane**: `src/taskpane/index.tsx` - Main React application entry point
- **Commands**: `src/commands/commands.ts` - Office ribbon command handlers

### Core Structure
- `src/taskpane/` - Main task pane React application
  - `components/App.tsx` - Root component with Fluent UI integration
  - `components/Header.tsx`, `components/HeroList.tsx`, `components/TextInsertion.tsx` - UI components
  - `taskpane.ts` - Excel API interaction functions (e.g., `insertText`)
- `src/commands/` - Office ribbon command implementations

### Key Technologies
- **React 18** with TypeScript for UI
- **Fluent UI React Components** for Microsoft design system
- **Office.js API** for Excel integration
- **Webpack 5** for bundling with HTTPS dev server
- **Babel** for TypeScript compilation

### Office Integration
- The add-in uses `manifest.json` for Office configuration
- Development server runs on `https://localhost:3000` with auto-generated certificates
- Production deployment requires updating the URL in webpack.config.js (currently set to contoso.com)

### Excel API Integration
- Excel operations are performed using `Excel.run()` async context
- Current implementation includes basic text insertion to cell ranges
- Functions in `taskpane.ts` handle direct Excel API calls

## Configuration Files

- `tsconfig.json` - TypeScript configuration targeting ES5 with React JSX
- `webpack.config.js` - Build configuration with separate entry points for taskpane and commands
- `manifest.json` - Office add-in manifest (note: there's also a `manifest.xml`)
- `babel.config.json` - Babel configuration for TypeScript and environment presets

## Development Notes

- The project uses Office add-in development certificates for HTTPS
- Hot module replacement is configured for React components
- The build process copies assets and transforms manifest URLs for production
- TypeScript is configured for React development with Office.js types
- ESLint uses Office add-in specific rules and React plugin