# 🪄 WindowsDictionary
📌 Automatically create shortcuts for your preferred web dictionaries and assign hotkeys to access them from any page or document.

---

## ✨ Overview

These files are lightweight Windows automation scripts written in VBScript.  
Each file, starting with "Dictionary", creates a **shortcut for itself**, prompts you to **assign a keyboard shortcut (Hotkey)**, and then safely moves the shortcut to your **Start Menu Programs** folder for easy access.

Ideal for developers, power users, and automation enthusiasts who want to trigger their scripts quickly from anywhere in Windows.

---

## 🎯 Purpose

It helps users quickly launch dictionary websites with a single key combination, improving workflow efficiency and reducing repetitive navigation.
It works with any open document — simply copy the text and press the assigned hotkey combination to activate it.

---

## 🚀 Features

- 📄 **Self-aware shortcut creation**  
  Automatically generates a `.lnk` file pointing to the running script.

- 🎹 **Hotkey assignment prompt**  
  Opens the Properties window so you can assign a global hotkey.

- 🗂️ **Start Menu integration**  
  Moves the shortcut into:

C:\Users<user>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs

---

## ⚠️ Restart Explorer workaround
- 🤦 Because Windows cannot always detect newly assigned hotkeys reliably, they may not work until Explorer is restarted.  
  I created **RestartExplorer.vbs** to safely reset Explorer.  
  ⚠️ **Warning:** This will close all your open windows, so use it with care.

---

## 🧩 How It Works

1. The script creates a shortcut for itself in the same folder.
2. If no hotkey is assigned, it opens the shortcut’s **Properties** window for you.
3. After you close Properties and confirm, the shortcut is moved to your Start Menu folder.
4. Optionally, Explorer restarts and key folders are reopened.
5. It uses the installed Chrome browser on your machine to open web dictionaries. Personally, I use Mozilla Firefox for my regular browsing, and Chrome exclusively for the dictionaries. If you’d like to switch browsers or use a different one, let me know by leaving a comment [here](https://github.com/hackus/WindowsDictionary/issues)

---

## 📖 Usage

1. Download or clone this repository.
2. Double-click any file starting with `Dictionary` to install it, for example: **`Dictionary<something>.vbs`**.
3. Assign a **Hotkey** when prompted. Once finished, either reboot your system or run [**RestartExplorer.vbs**](#-restart-explorer-workaround), making sure to follow the warnings.
4. After setup, you can copy any text from anywhere (Ctrl+C) and launch the script using your assigned hotkey.
5. If there are additional dictionaries you’d like, feel free to let me know and I’ll include them.

⚠️ Before clicking OK on the “Step 3: move shortcut to Start Menu Programs folder” message, make sure you’ve assigned a Hotkey in the Properties dialog.
If you accidentally close the Properties window without assigning one, just run the corresponding Uninstall script for that dictionary and reinstall it to try again.

⚠️ If you want to uninstall all the dictionaries at once run UninstallAllShortcuts.vbs

---

## 🧰 Requirements

- Windows 10 or later
- Windows Script Host (WScript/CScript) — enabled by default
- No external dependencies

---

## 🛠️ Optional Enhancements

You can customize
- The **destination folder** (e.g. Desktop, Quick Launch).
- Whether to **restart Explorer** automatically.

---

## 📂 Folder Structure

```
WindowsDictionary/
│
├── Common.vbs # Main script
├── Dictionary<...>.vbs # Script you should use.
├── Dictionary<...>_Uninstal.vbs # Script you should when you want to uninstall a dictionary.
├── ListHotkeys.vbs # Helps checking what hotkeys are currently assigned
├── RestartExplorer.vbs # Helps to restart the explorer
├── UninstallAllShortcuts.vbs # To uninstall all the dictionaries at once
├── README.md # Project documentation
└── .gitignore # Optional Git ignore file
```

---

## 🧑‍💻 Author

**Mircea Sirghi**  
Windows scripting & automation enthusiast  
💡 *“A single shortcut can save a hundred clicks.”*

---

## 🪐 License

This project is open-source under the **MIT License**.  
You’re free to use, modify, and distribute it with attribution.

---

### ⭐ If you find this project useful, give it a star on GitHub!
