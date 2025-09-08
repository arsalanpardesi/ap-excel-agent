## AP Excel Agent POC

AP Excel Agent POC is a proof-of-concept application that combines a web-based spreadsheet interface with a powerful AI agent to automate financial analysis. The system can ingest financial documents (like 10-K filings), parse them into structured data, and perform complex calculations based on natural language commands. It can also open Xlsx documents.

### Project Overview

The AP Excel Agent POC is an end-to-end system designed to demonstrate the power of agentic AI in the context of spreadsheet-based financial analysis.

**Key Features:**

    Interactive Spreadsheet UI: A modern, browser-based spreadsheet powered by Handsontable and the HyperFormula calculation engine.

    PDF Ingestion: Users can upload a 10-K PDF, which the server processes to extract financial statements (Income, Balance Sheet, Cash Flow).

    Xlsx Ingestion: User can open Excel Xlsx files.

    AI-Powered Parsing: Utilizes a local LLM via Ollama (qwen3:32b by default) to parse raw text from the PDF into structured JSON data and also has Gemini which can be called using your own API key.

    Agentic Analysis: A conversational agent that accepts natural language goals (e.g., "calculate gross margin %") and translates them into a sequence of precise, executable spreadsheet operations.

    Streaming & Debugging: The agent's thought process, from context analysis to final plan, is streamed live to a debug panel in the UI, providing transparency.

    Auditable & Reversible Actions: Every change to the workbook is logged as an atomic, reversible SheetOp, complete with an "undo" history.

    Data Provenance: Tracks the origin of data, allowing users to click a cell and see which document it was derived from.

    XLSX Export: The entire workbook can be exported as a standard .xlsx file.

### System Architecture

**Frontend (Client-Side)**

    UI Framework: A single, dependency-free HTML file (public/index.html) using vanilla JavaScript for simplicity and fast loading.

    Spreadsheet Grid: Handsontable is used to provide a feature-rich, Excel-like grid experience.

    Formula Engine: HyperFormula is integrated with Handsontable to handle real-time formula calculations directly in the browser.

    API Communication: The client interacts with the server via RESTful API endpoints and a Server-Sent Events (SSE) connection for real-time agent streaming.

**Backend (Server-Side)**

    Runtime: Node.js with TypeScript for type safety and modern JavaScript features.

    Web Server: Express.js handles API requests, serves the static frontend, and manages the agent's streaming endpoint.

    Core Logic (sheet.ts): The SheetModel class is the heart of the application. It maintains the workbook's state in memory and ensures all modifications are processed through an atomic, event-based system (SheetOps).

    PDF Processing (pdf.ts): Uses the pdfjs-dist library to extract raw text content from uploaded PDF files.

    AI Parser (parser.ts): Contains the logic for prompting a local LLM (via Ollama) to convert the extracted 10-K text into a structured TenKParsed JSON object. It includes robust error handling and a Zod schema for validation.

    AI Agent (agent.ts): Manages the core agent loop. It summarizes the workbook context, constructs a detailed prompt, and streams the LLM's response, which is then parsed into an executable plan.

### Getting Started

Follow these steps to set up and run the project locally.

**Prerequisites**

    Node.js: Version 18 or higher.

    Ollama: You must have Ollama installed and running.

    Ollama Model: Pull the required model. The default is qwen3:32b.
    Bash

    ollama pull qwen3:32b

**Installation & Setup**

    Clone the Repository:
    git clone https://github.com/arsalanpardesi/ap-excel-agent.git


**Install Dependencies:**
Navigate to the server directory and install the npm packages.

npm install

**Environment Variables:**
Create a .env file in the server directory by copying the example.
Bash

cp .env.example .env

Modify the .env file if your Ollama instance is not running on the default URL.
Code snippet

    # .env
    OLLAMA_BASE_URL=http://localhost:11434
    OLLAMA_MODEL=qwen3:32b
    PORT=3000

    GEMINI_API_KEY="YOUR API KEY HERE"

**Running the Application**

    Start the Server:
    From the server directory, run the development server.

    npm run dev

    Open in Browser:
    Navigate to http://localhost:3000 in your web browser.

You should now see the AP Excel Agent interface, ready for use.