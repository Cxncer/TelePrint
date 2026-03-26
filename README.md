
# 🖨️ Noble Printer Bot

Telegram bot for print shop automation. Users upload documents, select paper size, color, duplex, copies, and the bot prints to a network printer.

---

# 📁 Project Structure


noble-printer-bot/
  nobleprinter.py           # Main bot
  bot_tray_launcher.py      # System tray with log viewer
  start_bot_silent.vbs      # Startup script
  .env.example              # Environment template
  requirements.txt          # Dependencies

---

# 🚀 Quick Start

  1. Install Python 3.12 and SumatraPDF
- Python: [python.org](https://www.python.org/downloads/)
- SumatraPDF: [sumatrapdfreader.org](https://www.sumatrapdfreader.org/download-free-pdf-viewer)

  2. Clone and setup

git clone https://github.com/cxncer/noble-printer-bot.git
cd noble-printer-bot
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt


  3. Configure
Copy `.env.example` to `.env` and edit:

BOT_TOKEN=your_bot_token
STAFF_GROUP_ID=-1001234567890
PRINTER_IP=192.168.1.100
PRINTER_NAME=iR-ADV C3530(2)


  4. Run
- Test: `python nobleprinter.py`
- Background: double-click `start_bot_silent.vbs`

---

# 🖥️ Commands

| Command | Description |
|---------|-------------|
| `/ping` | Check if bot is alive |
| `/debug` | Show debug info |
| `/quick_test` | Quick print test (reply to document) |

---

# 📊 Pricing

- BW A4: 200 riels/page
- Color A4: 300 riels/page
- A3: 2× base
- Double-sided A4: +25%

---

# 🔧 Troubleshooting

- **Bot not responding**: Check token and group ID.
- **Printer not printing**: Verify printer IP and name.
- **Duplex not working**: Ensure printer supports duplex and check Windows driver settings.

---

# 🙏 Credits

- [python-telegram-bot](https://github.com/python-telegram-bot/python-telegram-bot)
- [PyPDF](https://pypi.org/project/pypdf/)
- [PyStray](https://github.com/moses-palmer/pystray)

Special thanks to **[DeepSeek](https://deepseek.com)** for development assistance, code optimization, and troubleshooting.

---

*Made with ❤️ in Cambodia 🇰🇭*
