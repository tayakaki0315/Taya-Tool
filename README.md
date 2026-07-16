# Taya Expert List Builder

A browser-based tool for turning pasted expert profile text into a structured Excel expert list.

## Features

- Parses multiple expert profiles in one paste
- Extracts introductions, screening Q&A, employment history, fees, and availability
- Lets you review and edit every field before export
- Supports one sheet, one sheet per expert, or drag-and-drop custom sheet grouping
- Generates a formatted `.xlsx` file entirely in the browser
- Keeps pasted expert data local to the browser

## Live site

https://taya-expert-list.tayakaki0315.chatgpt.site

## Local development

Requirements: Node.js 22.13 or later.

```bash
npm install
npm run dev
```

Production validation:

```bash
npm test
```

## Main files

- `app/page.tsx`: parser, editor, sheet organizer, and Excel export
- `app/globals.css`: application styling
- `package.json`: scripts and dependencies

## Privacy

The application processes pasted expert information in the browser and does not upload or store it on a server.
