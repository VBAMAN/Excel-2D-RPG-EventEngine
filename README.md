# Excel 2D RPG EventEngine

**Version: 0.5.1**

Part of the **Excel 2D RPG** series.

Inspired by classic JRPG event systems.

A worksheet-based event engine for building simple RPG systems in Excel using VBA.

---

## 🎮 Demo

https://youtu.be/rsMcP9zigJc

---

## ✨ Features

- Event-driven system (STEP / COMMAND / etc.)
- Worksheet-based map control
- Direction-aware triggers
- Simple and extensible VBA structure

---

## 🕹 Controls

- Arrow Keys: Move player
- X Key: Execute command (TALK)

---

## 🧩 Event Example

```vba
Set ev = New EVENT_HANDLE
ev.eventID = "EV_CMD"
ev.eventMap = 1
ev.EventTriggerType = "COMMAND"
ev.TriggerCommand = "TALK"
ev.IsEnabled = True
ev.IsCompleted = False
ev.Action = "DisplayMessage"
ev.eventName = "Command Window"
```

## 🏗 Structure

- `EVENT_HANDLE` : Event data class
- `ETRG_10_init` : Event initialization
- `ETRG_20_main` : Event execution logic
- `ETRG_15_validate` : Event validation (debug)
- `PLAYER_*` : Player control & movement

## 🔗 Related Project

Excel-2D-RPG-Core (UserForm-based core system)

## 📌 Notes

This project focuses on event handling logic.
Designed for learning and experimentation with Excel-based game systems.


## 🛠 Requirements

Microsoft Excel (VBA enabled)

## 📁 How to Run

1. Open `excel_rpg_event_engine_demo.xlsm` in Excel
2. Enable macros when prompted
3. Click the **"START"** cell on the worksheet to begin
4. Click the **"STOP"** cell to end the game

