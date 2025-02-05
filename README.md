An attempt to bring the core abilities of TAPI into modern programming languages, so that everyone can use them as simple single files.

## Background
**Microsoft TAPI (Telephony Application Programming Interface)** is unified modem control API for Windows, which allows making calls, displaying calls (including caller ID), etc. regardless of the modem.
<br />Prior to TAPI, each modem used different commands to pull each action.
<br />In addition, TAPI allows using modems in an asynchronous (parallel) way, so that the same modem can be used at the same time by different programs.

Unfortunately, it's hard to find new programs or codes that use TAPI, especially free ones. It might be because Microsoft's TAPI documenation is super hard to deicpher.

## call.py
A Python script that can call a numbers, which is delivered via a command line parameter.

**Usage:** `call.py [phone number]` (without brackets)

**Configuration:**
1. Change `modemAll = False` to `modemAll = True` to make the call in every possile modem (if you have more than one modem).
1. Change the `USB` in `modem = "USB"` to another word that is contained in the modem's name (so it will know which modem to use).

## callerid.py
A Python script that pops up incoming calls including caller ID. **This one is exclusive** - try searching for anything similar online and you'll only find unanswered questions or non working code <sub>~~[(for example)](https://github.com/firstoxe/TAPI-Event-monitor/issues/1)~~</sub>.

**Usage:** just run the script as-is.

**Configuration:**
1. Change `print_console = False` to `print_console = True` if you prefer the output to run in the OS console instead of a popup GUI.
1. Change `modemAll = True` to `modemAll = False` to not make the caller ID work in every possile modem (if you have more than one modem).
1. Change the `USB` in `modem = "USB"` to another word that is contained in the modem's name (so it will know which modem to use for Caller ID).
