<div align="center">

# 📊 BA-Dashboard

**A Business Analyst Productivity Hub for managing tasks, generating documents, and visualizing data — all in one place.**

[![Node.js](https://img.shields.io/badge/Node.js-18+-339933?style=for-the-badge&logo=node.js&logoColor=white)](https://nodejs.org)
[![Express.js](https://img.shields.io/badge/Express.js-4-000000?style=for-the-badge&logo=express&logoColor=white)](https://expressjs.com)
[![SQLite](https://img.shields.io/badge/SQLite-003B57?style=for-the-badge&logo=sqlite&logoColor=white)](https://www.sqlite.org)
[![JavaScript](https://img.shields.io/badge/JavaScript-ES6+-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)](https://developer.mozilla.org/en-US/docs/Web/JavaScript)
[![Render](https://img.shields.io/badge/Deployed%20on-Render-46E3B7?style=for-the-badge&logo=render&logoColor=white)](https://render.com)

</div>

---

## Table of Contents

- [Introduction](#introduction)
- [Installation](#installation)
- [Features](#features)
- [Technologies](#technologies)
- [API Reference](#api-reference)
- [Contributing](#contributing)
- [License](#license)
- [Acknowledgements](#acknowledgements)

---

## Introduction

BA-Dashboard is a full-stack productivity hub built specifically for Business Analysts. It provides a clean, lightweight interface to manage day-to-day BA tasks such as uploading and processing files, generating Word documents, visualizing data, and integrating with enterprise systems using NTLM authentication.

The frontend is built with plain HTML, CSS, and JavaScript — keeping it fast and dependency-free. The backend is powered by Node.js and Express.js, with SQLite as a lightweight embedded database. The application is deployed on Render.com and supports enterprise-grade authentication through NTLM for SharePoint and internal network integration.

BA-Dashboard/
├── frontend/     →  Vanilla HTML, CSS, JavaScript  (port 5173)
├── backend/      →  Node.js + Express REST API      (port 3000)
├── render.yaml   →  Render deployment configuration
└── .gitignore



---

## Installation

To set up and run the project locally, follow these steps:

**1. Clone the repository:**

```bash
git clone https://github.com/AdityaS326/BA-dashboard.git
cd BA-dashboard
2. Install backend dependencies:


cd backend
npm install
3. Set up environment variables:

Create a .env file inside the backend/ folder:


cd backend
cp .env.example .env
Fill in the required values in backend/.env:


PORT=3000
NODE_ENV=development

# NTLM credentials for enterprise/SharePoint integration
NTLM_USERNAME=your-domain-username
NTLM_PASSWORD=your-domain-password
NTLM_DOMAIN=your-domain
NTLM_URL=http://your-sharepoint-url
4. Start the backend server:


# Development (with auto-reload)
cd backend
npm run dev

# Production
cd backend
npm start
5. Start the frontend:

Open a new terminal:


cd frontend

# Install frontend dev dependencies (first time only)
npm install

# Development
npm run dev

# Production preview
npm start
Service	URL
Frontend	http://localhost:5173
Backend API	http://localhost:3000
Features
Visual data representation: A clean dashboard interface that displays BA metrics and data in an easy-to-understand visual format without the overhead of a heavy frontend framework.

Word document generation: Automatically generate structured .docx Word documents directly from the dashboard, saving time on repetitive reporting tasks.

File uploads: Upload and manage files through the interface using Multer-powered endpoints, with support for various file types needed in BA workflows.

Media processing: Process and handle media files using ffmpeg integration, enabling audio and video handling within the productivity hub.

Enterprise authentication: Connect to internal enterprise systems and SharePoint using NTLM authentication, making it compatible with corporate network environments.

Lightweight SQLite database: All data is stored in an embedded SQLite database — no separate database server required, making setup and deployment simple.

REST API backend: A well-structured Express.js API handles all business logic, file handling, document generation, and external integrations.

Render deployment: The application is configured for seamless deployment on Render.com via render.yaml with minimal setup.

Technologies
HTML, CSS, JavaScript: The frontend is built with vanilla web technologies — no framework overhead, fast loading, and easy to maintain.

Node.js: A JavaScript runtime built on Chrome's V8 engine, used to run the backend server with high performance and non-blocking I/O.

Express.js: A minimal and flexible Node.js web framework that provides a robust set of features for building the REST API.

SQLite (better-sqlite3): A self-contained, serverless, embedded SQL database engine. Ideal for lightweight applications that don't require a separate database server.

Multer: A Node.js middleware for handling multipart/form-data, used for file uploads within the application.

docx: A JavaScript library for creating and generating .docx Word documents programmatically from templates or dynamic data.

ffmpeg-static: A static build of FFmpeg bundled with Node.js, used for processing audio and video files within the backend.

httpntlm: A Node.js library that handles NTLM authentication, enabling the application to connect securely to Microsoft enterprise systems and SharePoint.

node-fetch: A lightweight module that brings the browser fetch API to Node.js, used for making HTTP requests to external services.

dotenv: A zero-dependency module that loads environment variables from a .env file into process.env, keeping credentials out of the codebase.

nodemon: A development utility that automatically restarts the Node.js server whenever file changes are detected, speeding up the development workflow.

live-server: A lightweight development server with live reload capability for the frontend, enabling instant browser refresh on file changes.

serve: A simple static file server used to serve the frontend in production and preview environments.

Render.com: A cloud platform used to deploy and host both the frontend and backend, configured through render.yaml.

API Reference
Base URL: http://localhost:3000

Method	Endpoint	Description
GET	/health	Health check — returns server status
POST	/api/upload	Upload a file via multipart form
GET	/api/documents	List all generated documents
POST	/api/documents/generate	Generate a Word .docx document
GET	/api/data	Fetch dashboard data from SQLite
POST	/api/auth/ntlm	Authenticate with enterprise NTLM
Contributing
Contributions are welcome! Please follow these steps to contribute:

Fork the repository.
Create a new branch: git checkout -b feature/your-feature-name
Make your changes.
Commit your changes: git commit -m 'Add your feature description'
Push to the branch: git push origin feature/your-feature-name
Open a Pull Request.
License
This project is licensed under the MIT License. See the LICENSE file for details.

Acknowledgements
Express.js — For providing a fast and minimal Node.js web framework for the backend API.
better-sqlite3 — For making SQLite integration in Node.js simple, synchronous, and performant.
docx — For enabling programmatic Word document generation directly from Node.js.
Multer — For seamless file upload handling in the Express backend.
httpntlm — For making enterprise NTLM authentication accessible in a Node.js environment.
ffmpeg-static — For bundling FFmpeg with the application for easy media processing.
Render.com — For providing a straightforward and reliable platform for deployment.
