# Age of Empires IV Tournament Platform: Dashboard & Bot Integration

## Overview

This project is a custom-built tournament management system designed for Age of Empires IV. It combines a responsive web dashboard with a backend Telegram bot to streamline competitive gaming operations[cite: 3]. The system handles player registrations, match tracking, announcements, and betting logs in real time.

By integrating automated messaging through Telegram with a dynamic web interface, the platform eliminates manual administrative overhead and provides players with an interactive, centralized hub for tournament data.

* **Web-App Link:** (https://script.google.com/macros/s/AKfycbzIvh3JCWQ8Z8wPdXMaKR4Pbu7TFF6AIpAGy7PelBt0yidYKB4_ceTup_XN4Jy-GjO-8g/exec)
---

## Tech Stack

**Frontend**
* HTML5 & CSS3
* JavaScript (ES6+)
* Responsive UI Design

**Backend & Integration**
* Telegram API (Webhooks & Bot Operations)
* Google Apps Script (Data routing and backend logic)
* JSON (Local database structuring)

**Version Control**
* Git
* GitHub

---

## Data Architecture

**Source**
A modular, lightweight architecture prioritizing fast updates and real-time communication.

* **Player Database:** Local JSON structures storing player sign-ups, current brackets, and match statistics.
* **Betting Logs:** Integrated tracking for match predictions and wager logging.
* **Webhook Routing:** Apps Script endpoints handling payloads between the web frontend and the Telegram bot API.

---

# Project Workflow

## Stage 1 — Web Interface Development (HTML/CSS/JS)
The foundation required building a fast, responsive dashboard for players to view tournament standings and match schedules.
* Designed a clean UI customized for gaming analytics and bracket visualizations.
* Implemented JavaScript to dynamically render match data and player statistics on the frontend.
* Formatted all frontend assets with standard `.html` and `.js` extensions for clean hosting.

## Stage 2 — Telegram Bot Integration
To automate tournament communications, a dedicated Telegram bot was developed as the primary operational engine.
* Built webhook handlers to process commands for player sign-ups and match reporting.
* Automated real-time match announcements directly into the tournament Telegram channel.
* Integrated a logging system to record and calculate user betting/predictions on match outcomes.

## Stage 3 — Backend Synchronization & Apps Script
Bridged the gap between the static dashboard and the active Telegram bot using Google Apps Script.
* Configured the script to listen for incoming Telegram webhooks and route data appropriately.
* Ensured match results reported via Telegram instantly updated the local data structures feeding the dashboard.

---

# Key Application Features

### Automated Match Announcements
The Telegram bot acts as an automated tournament admin, instantly broadcasting match starts, results, and bracket updates to all participants without manual input.

### Integrated Betting & Prediction Logging
Allows community members to engage with the tournament by logging match predictions, automatically tracking success rates and betting data through the backend.

### Dual-System Architecture
Combines the accessibility of a web dashboard for data visualization with the real-time push notifications of a Telegram bot, ensuring players never miss a match.

---

# Repository Structure

```text
project-root
│
├── 00_javascripts               # Frontend JS scripts[cite: 5]
│   └── (Bot code ignored via .gitignore)
│
├── 01_html_scripts              # Frontend UI pages and layouts[cite: 5]
│
├── 02_local_database            # Ignored via .gitignore (Contains active player data)[cite: 5]
│
├── 03_backups                   # Ignored via .gitignore (Historical states and scripts) 
│
├── .gitignore                   # Security and version control exclusions[cite: 4, 5]
│
└── README.md                    # Project documentation[cite: 4, 5]
```
---

# Lessons Learned

The primary motivation behind this project was solving the logistical chaos of organizing competitive gaming tournaments. Managing sign-ups, reporting scores, and keeping players updated manually across different platforms often leads to delays and confusion.

Key takeaways from this project:

* Event-Driven Architecture: Utilizing Telegram webhooks taught the importance of event-driven data flow, ensuring the system only uses resources when a command or match update is triggered.
* Community-Centric UI: Building a dashboard for gamers required focusing on clear, immediate data visualization (brackets, stats, betting logs) over text-heavy interfaces.
* Having an active/live database of players to practice on is a huge help when creating a dataset as opposed to having to do it with assumptions and self driven tests.
* The more granualr the application of an idea is, the further the developer understands the importance of an advanced designed database.
* prior to this project, my work always consisted of multiple lines of code as well as a readable view inside of spreadsheet.
But after working with AppScript and witnessing the level of maneuver it gives you, I am quite fond of it! 

---

# Future Improvements

Potential extensions include:

* Automated Bracket Generation: Implementing algorithms to automatically seed and generate double-elimination brackets based on initial sign-ups.
* Adding a Team Tournament Generaor: Implemeting that option is just a matter of demand now!
* Running fyrther analytical reports on civ stats and other features of the games such as strategies used and more.
* Making the web-app more global and inviting a bigger player-base to compete with it!
* Perhaps using this experience to design a web-app that can handle tournament design for popular games as well.
