# Racal-3271-Control

A modern (kinda) Windows GUI application for controlling a Racal 3271 RF Signal Generator installed in an Agilent E8403A VXI mainframe, using an Agilent E1406A Command Module, via GPIB / NI-VISA.

This project provides a clean, engineer-friendly interface for frequency, level, RF output control, presets, and manual SCPI-style command entry.

---

## Hardware Setup

This software is designed for the following configuration:

- Racal 3271 RF Signal Generator (9 kHz – 2.4 GHz)
- Agilent E8403A VXI Mainframe
- Agilent E1406A VXI Command Module
- GPIB connection to PC (NI, Agilent/Keysight, or compatible)
- NI-VISA installed on Windows

Typical VISA resource string:

```
GPIB0::10::12::INSTR
```

Where:
- 10 = E1406A primary GPIB address
- 12 = Racal 3271 secondary address (VXI logical address)

---

## Features

- Native NI-VISA communication
- WinForms UI with Light and Dark themes
- Frequency control with units (Hz, kHz, MHz, GHz)
- RF level control (dBm)
- RF output ON / OFF control
- Mode selection
- Presets / profiles for common RF setups
- Manual SCPI command entry with live responses (acting as a terminal)
- Session log with timestamps
- Automatic VISA resource scanning
- Safe connect / disconnect handling

---

## Getting Started

### 1. Install Prerequisites

- Windows 11
- NI-VISA 

Download NI-VISA from:
https://www.ni.com/en/support/downloads/drivers/download.ni-visa.html#575764

- Visual Studio 2026

---

### 2. Clone the Repository

```
git clone https://github.com/tomv2/Racal-3271-Control.git
```

Open the solution in Visual Studio.

---

### 3. Build and Run

- Set build configuration to: Any CPU
- Run the application
- Select the correct VISA resource
- Click Connect

The application will query:

```
*IDN?
```

Expected response example:

```
Racal Instruments,3271,202307/882,44533/445/02.15
```

---

## VXI Addressing and DIP Switch Configuration (Important)

The Racal 3271 uses an on-board DIP switch to set its **VXI logical address** (secondary address).

### Known Working Configuration

This project is tested using:

- **VXI logical address: 12**
- VISA resource: `GPIB0::10::12::INSTR`

The DIP switch is configured to represent logical address **12 (decimal)**.

This address is:
- Within the valid VXI logical address range
- Properly detected by the Agilent E1406A as a static device
- Accessible via NI-VISA message-based communication

### Address 255 Warning (Do Not Use)

Using a logical address of **255** did **not** work in this setup.

In VXI systems:
- Logical address **255** is likely reserved (I didn't actually read the manual and check...) 
- It should not be assigned to normal instruments
- Some controllers may partially enumerate the device but refuse access

I also tried a couple of other addresses around 255 and no success with them either.

---

## Presets

Built-in presets include examples such as:

- 1 GHz @ -20 dBm
- 2 GHz @ -20 dBm
- GPS L1 (1.57542 GHz)
- WiFi Channel 1 (2.412 GHz)
- RF OFF (safe state)

Presets only load values into the UI.  
Click Apply to send the preset values to the instrument.

---

## Manual Command Entry

The manual command box supports any valid Racal 3271 command.

Example commands:

```
*IDN?
CFRQ?
RFLV?
CFRQ 2GHZ
RFLV -10DBM
RFLV:OFF
```

- Commands ending in `?` are queried automatically
- Responses are logged in the session log

---

## Notes and Limitations

- SWEEP and LIST modes are sent as best-effort commands
- It's possible not all Racal 3271 firmware versions support every SCPI-style command
- VXI logical address must be set correctly via DIP switches on the Racal module
- Ensure the E1406A detects the Racal as a device during boot (can be done via RS232 terminal)

---

## Why This Project Exists

Modern PCs no longer support legacy VXI control software well from what I found. There was no plug and play code or software to run and get this setup running.

I had never used VXI systems before either and so it was a bit of learning to understand how they work and how to get results.

This project provides a lightweight, modern, maintainable replacement for controlling classic RF instrumentation without relying on obsolete environments.

All you rely on is the NI drivers.

---

## Further Work

Contributions are welcome, especially for:

- Additional presets
- Sweep and list UI improvements
- Instrument state polling (optional setting)
- Exportable profiles
- Better looking UI
- Option to import command scripts (useful for calibration automation)

---

## Contact

Feel free to open an issue and I'll try help!
