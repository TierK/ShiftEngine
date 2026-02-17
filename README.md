# ShiftEngine

**ShiftEngine** is a robust automation solution for managing 14-day shift cycles within Google Workspace. Designed for high-reliability environments, it handles everything from constraint collection to algorithmic assignment and dark-mode analytics.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Developer: TierK](https://img.shields.io/badge/Developer-TierK-blue)](https://github.com/TierK)

## 📌 Latest Version
- **Current version:** `v0.16.1`
- Synced with Apps Script deployment `v0.16.1`.
- If GitHub still shows older content, refresh branch and pull latest commits.

---

## 🛠 Setup & Access
> ⚠️ **Note on SPREADSHEET_ID**: For security reasons, the Master Spreadsheet ID is not public. 
> To get the template and your unique ID, please [**Contact the Developer**](mailto:your-email@example.com?subject=ShiftEngine%20Access%20Request).

---

## 📸 System Overview

### 1. Management & Target Tracking
The system synchronizes with the `Staff` sheet to monitor shift targets vs. actual assignments.
![Config Staff](screenshots/Staff.png)

### 2. Main Scheduling Interface
A 14-day grid featuring real-time validation logic and status indicators.
![Main Schedule](screenshots/Schedule.png)

*Automated color coding for Week 1 and Week 2:*
![Weekly Detail](screenshots/Export.png)

### 3. Integrated Command Center (`פקודות✨`)
Custom UI menu providing direct access to the engine's core functions.
![Custom Menu](screenshots/commands.png)

---

## 📝 Google Form & Response Structure

For the engine to process constraints correctly, the linked Google Form must follow a strict structure. The system maps the **User Interface (Form)** directly to the **Data Ingress (Sheet)**.

### 1. The Interface (Frontend)
* **Employee Name:** Must be a Dropdown or Short Answer that **exactly** matches the `Staff` sheet.
* **Grid/Checkboxes:** 14 separate questions for the 2-week cycle.
* **Allowed Values:** `בוקר` (Morning), `צהריים` (Afternoon), `לילה` (Night).

![Google Form UI](screenshots/form.png)

### 2. The Data Ingress (Backend)
The engine uses `cleanupDuplicateResponses()` to ensure that if an employee submits the form multiple times, only their **latest** submission is kept.

![Form Responses](screenshots/Responses.png)
*Above: How the `Responses` sheet looks after successful synchronization.*

---

## 🛠 Technical Features

* **Smart Algorithm:** Prioritizes "Rare" availability and balances based on defined targets.
* **Invisible Markers:** Uses `\u200B` (Zero-width space) to distinguish between automated and manual entries.
* **Conflict Engine:** Detects double shifts and "Night-to-Morning" violations (rest time protection).
* **Dynamic Export:** Creates a standalone file with a unique color theme for every new export to maintain visual variety.
* **Dark Mode Analytics:** Full-scale dashboard including:
    * Workload Distribution (Pie Chart)
    * Night Shift Analytics (Purple Theme)
    * Shabbat Distribution (Gold Theme)

![Analytics Dashboard](screenshots/analytic.jpg)

---

## ⚙️ DevOps & Deployment

### Required Triggers
For the automation to function correctly, the spreadsheet owner must manually set up the following triggers in the Apps Script console:

![Triggers Setup](screenshots/triggers.png)

| Function | Event Source | Event Type | Description |
| :--- | :--- | :--- | :--- |
| `applyGreenIfAllOk` | From Spreadsheet | On form submit | Processes new constraints from Google Forms. |
| `installedOnEdit` | From Spreadsheet | On edit | Updates UI colors and conflict validation instantly. |

### Global Configuration
The system relies on a central `CONFIG` object. Note that the ID is abstracted for security.

```javascript
const CONFIG = {
  VERSION: "0.16.1",
  /** * SPREADSHEET_ID and FORM_URL are hidden for security. 
   * To request access or a template of SPREADSHEET or/end GOOGLE_FORM, contact: kimbfsd@gmail.com
   */
  SPREADSHEET_ID:'YOUR_ SPREADSHEET_ID',
  FORM_URL: 'YOUR_FORM_URL'
  // ... rest of config
};
````
## 📄 License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 🗂 Local Module Structure (for clasp workflow)

The codebase is now split into logical folders locally:

- `core/`
- `triggers/`
- `services/`
- `ui/`
- `utils/`
- `main.gs`

### How to see this in Apps Script editor
Google Apps Script editor does **not** preserve local folder hierarchy as real folders in the UI. It displays files as a flat list.

To work with this structure:
1. Develop locally in this repo with folders.
2. Sync with Apps Script via clasp (`clasp push`).
3. In the Apps Script web editor you will see the files, but without folder tree nesting.

If you want clear grouping inside Apps Script UI, use filename prefixes (for example: `services.schedule.service.gs`, `ui.menu.ui.gs`) — this project already uses that style.


## ▶️ Как запустить и посмотреть локально (до публикации)

Ниже быстрый путь, чтобы сначала проверить всё локально, а потом отправить в Apps Script:

1. **Установить Node.js** (если ещё не установлен).
2. **Установить clasp**:
   ```bash
   npm i -g @google/clasp
   ```
3. **Авторизоваться в Google**:
   ```bash
   clasp login
   ```
4. **Привязать проект к вашему Apps Script** (в корне проекта создать `.clasp.json`):
   ```json
   {
     "scriptId": "ВАШ_SCRIPT_ID",
     "rootDir": "."
   }
   ```
   `scriptId` берётся из Apps Script Project Settings.
5. **Локальная проверка синтаксиса** (без деплоя):
   ```bash
   node -e 'const fs=require("fs");const cp=require("child_process");const files=cp.execSync("find core triggers services ui utils -name \"*.gs\" | sort").toString().trim().split("\n");for(const f of files){new Function(fs.readFileSync(f,"utf8"));}console.log("OK",files.length)'
   ```
6. **Отправить в Apps Script**:
   ```bash
   clasp push
   ```
7. **Открыть в браузере редактор Apps Script**:
   ```bash
   clasp open
   ```

После `clasp push` вы увидите все `.gs` файлы в редакторе Apps Script (плоским списком, без реальных папок).
