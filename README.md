# 🚀 Enterprise WhatsApp Birthday Automation & CRM Audit Engine

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![WhatsApp API](https://img.shields.io/badge/WhatsApp%20Business%20API-25D366?style=for-the-badge&logo=whatsapp&logoColor=white)
![Atom Engine](https://img.shields.io/badge/Atom%20Flowbuilder-FF5B00?style=for-the-badge&logo=atom&logoColor=white)
![RevOps](https://img.shields.io/badge/RevOps%20Architecture-2C3E50?style=for-the-badge&logo=google-cloud&logoColor=white)

**PROFILE:** Michael Sancivier — Technical Onboarding & CSM Specialist | RevOps Process Strategist  
**CORE SKILLS:** Conversational Architecture | FinOps Strategy | API Webhooks | CRM Data Pipelines | Agile Execution

---

## 💡 System Overview

An enterprise-grade, autonomous CRM engine designed to ingest massive customer databases (+20,000 records), auto-sanitize international phone formats (E.164), apply anti-duplication cost shields, execute personalized WhatsApp HSM birthday greetings via Webhooks, and generate daily executive HTML audit reports.

Built for Shared Financial Services (CIDRE IFD), this decoupled system eliminates human manual intervention while safeguarding Meta HSM template costs and protecting credit officers' SLA.

---

## 📊 1. Business Impact & System Metrics

| Strategic Dimension | Execution / Technical Impact |
| :--- | :--- |
| **Data Ingestion Capability** | **100% Autonomous Ingestion** from Google Drive supporting `.xlsx` and `.csv` files. |
| **Meta HSM FinOps Protection** | **Anti-Duplication Shield** via persistent Script Properties, preventing double-firing costs and Meta spam flagging. |
| **Agent SLA Shielding** | **0% Human Inbox Clutter** through automated conversation routing and auto-closure (`RB - Resuelto por Bot`). |
| **CRM Hygiene & Normalization** | **E.164 Universal Phone Masking** automatically converting raw local inputs into international formats (`+591`). |
| **Executive Governance** | **Daily HTML Audit Dispatch** delivering real-time database health, ingested files history, and status breakdown. |

---

## 🏗️ 2. System Architecture & Data Flow

[Google Drive Input Folder]
│
▼ (Daily Cron Trigger - Apps Script)
[Auto-Ingestion & E.164 Sanitization]
│
▼
[Anti-Duplication Check (ScriptProperties)]
│
├──► [Match Found] ──► Log as Omitted (Yellow Status)
│
└──► [New Record]  ──► [Atom Webhook API Trigger (POST)]
│
├──► Dispatches WhatsApp HSM ('envios_cumpleanos_cidre')
├──► Auto-closes thread ('RB - Resuelto por Bot')
└──► Dispatches Daily Executive HTML Audit Email

---

## ⚙️ 3. Key Technical Capabilities

1. **Decoupled Configuration Object (`CONFIG`):** Allows instant reusability across any client or enterprise workspace by updating environment variables.
2. **E.164 International Sanitization Engine:** Handles variable digit lengths and injects country codes dynamically (`591` Bolivia, `55` Brazil, etc.).
3. **Persistent Deduplication Lock:** Uses Google Apps Script `PropertiesService` to maintain a memory lock for 24-hour windows.
4. **Resilient Drive File Management:** Automates file state transitions from `01_Input` to `02_Processed` post-execution with full error trapping.
5. **Branded HTML Executive Reporter:** Compiles daily execution metrics, processed source files, and individual transaction statuses into a responsive HTML email layout.

---

## 🛠️ 4. Configuration & Deployment

### Global Parameterization (`appscript_master.js`)

```javascript
const CONFIG = {
  CLIENTE_NOMBRE: "CIDRE IFD",
  EMAIL_CONTACTO_CLIENTE: "Katia",
  EMAILS_NOTIFICACION: "michael.sancivier@atomchat.io, khuici@cidre.org.bo",
  
  // Webhook Integration
  ATOM_WEBHOOK_URL: "[https://api.atomchat.io/v1/webhooks/hsm/trigger](https://api.atomchat.io/v1/webhooks/hsm/trigger)",
  ATOM_TOKEN: "YOUR_BEARER_TOKEN_HERE",
  
  // Drive Repositories
  FOLDER_ENTRADA_ID: "YOUR_INPUT_FOLDER_ID",
  FOLDER_PROCESADOS_ID: "YOUR_PROCESSED_FOLDER_ID",
  
  // E.164 Telephony Standard
  CODIGO_PAIS_DEFAULT: "591",
  LARGO_NUMERO_LOCAL: 8,
  TIMEZONE: "GMT-4"
};
```

## 📂 5. Repository Structure

1. appscript_master.js: Complete Google Apps Script backend containing drive processing, API webhooks, deduplication logic, and HTML email generator.

2. README.md: System documentation, business architecture, and setup guide.

## 👤 Author
### Michael Sancivier
### Technical Onboarding & CSM Specialist | RevOps Process Strategist
