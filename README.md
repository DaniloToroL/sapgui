# SAPGUI - SAP Scripting in Python

This module is to be able to automate SAP in a "simpler" way with python. It allows to use for example ```getattr``` without generating an error and it is expected to have integration with rfc

Running on windows and any pull request is welcome

## Usage

```
from sapgui.sapgui import SAPGui

sap_object = SAPGui()
sap_object.open_connection("conn", timeout=10, start_app=True)
sap_object.login("user", "12345")

object = sap_object.get_object("wnd[0]/usr/tabsTABSPR1/tabpSP02")
object.select()
sap_object.sendKey("ctrl+f")
sap_object.sendKey(vkey=0)
```


## SAP GUI Scripting

### Related links
- [SAP Documentation](https://help.sap.com/doc/9215986e54174174854b0af6bb14305a/760.01/en-US/sap_gui_scripting_api_761.pdf)
- [Tracker](https://tracker.stschnell.de/)

- [Virtual Keys](https://experience.sap.com/files/guidelines/References/nv_fkeys_ref2_e.htm)

- [SAP shortcut](shortcut:https://www.wcupa.edu/_Information/AFA/SAP/Shortcut_Keys.pdf)


