# ShiftEngine v0.14

Automated scheduling engine for bi-weekly shift management with real-time validation and dynamic reporting.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Developer: TierK](https://img.shields.io/badge/Developer-TierK-blue)](https://github.com/TierK)

## 📸 System Preview

### 1. Management Dashboard & Metrics
The system tracks target shifts versus actual assignments for each employee.
![Config Staff](screenshots/image_b54765.png)

### 2. Main Scheduling Interface
The 14-day interactive grid where the magic happens.
![Main Schedule](screenshots/image_b543df.png)
*Detailed view of the weekly distribution:*
![Weekly Preview](screenshots/Screenshot%202026-02-04%20105355.png)

### 3. Data Collection
Constraints are collected via Google Forms and automatically synced.
![Form Responses](screenshots/image_b540dc.png)

### 4. Custom Commands (`פקודות✨`)
Built-in menu for administrative tasks.
![Custom Menu](screenshots/image_b5481c.png)

---

## ⚙️ Trigger Configuration (Critical Step)

For the automation to work, the owner must manually set up these triggers in the Apps Script dashboard:
![Triggers Setup](screenshots/image_b52e19.png)

| Function | Event Source | Event Type |
| :--- | :--- | :--- |
| `applyGreenIfAllOk` | From Spreadsheet | On form submit |
| `installedOnEdit` | From Spreadsheet | On edit |

---

## 📊 Analytics & Export
Every exported file includes a dedicated dashboard with visual workload analysis.
![Analytics Dashboard](screenshots/image_b5a1b8.png)

## 📄 License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
