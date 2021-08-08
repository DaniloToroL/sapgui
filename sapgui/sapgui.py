import os
import subprocess
import time
from typing import Union

from win32com import client

class SapVKey:
    """
    See all vkey in C:\Program Files (x86)\SAP\Frontend\SAPgui\SAPguihelp\SAPGUIScripting.chm
    """

    CtrlA = 72
    CtrlB = 87
    CtrlC = 92
    CtrlD = 73
    CtrlE = 70
    CtrlF = 71
    CtrlF1 =25
    CtrlF2 =26
    CtrlF3 =27
    CtrlF4 =28
    CtrlF5 =29
    CtrlF6 =30
    CtrlF7 =31
    CtrlF8 =32
    CtrlF9 =33
    CtrlF10 =34
    CtrlF11 =35
    CtrlF12 =36
    CtrlG = 84
    CtrlIns = 77
    CtrlK = 88
    CtrlN = 74
    CtrlO = 75
    CtrlP = 86
    CtrlR = 85
    CtrlShift0 = 22
    CtrlShiftF1 = 37
    CtrlShiftF2 = 38
    CtrlShiftF3 = 39
    CtrlShiftF4 = 40
    CtrlShiftF5 = 41
    CtrlShiftF6 = 42
    CtrlShiftF8 = 44
    CtrlShiftF9 = 45
    CtrlShiftF10 = 46
    CtrlShiftF11 = 47
    CtrlShiftF12 = 48
    CtrlT = 89
    CtrlV = 93
    CtrlX = 91
    CtrlY = 90
    Enter = 0
    F1 = 1
    F2 = 2
    F3 = 3
    F4 = 4
    F5 = 5
    F6 = 6
    F7 = 7
    F8 = 8
    F9 = 9
    F10 = 10
    F11 = 11
    F12 = 12
    Help = 13
    ShiftF2 = 14
    ShiftF3 = 15
    ShiftF4 = 16
    ShiftF5 = 17
    ShiftF6 = 18
    ShiftF7 = 19
    ShiftF8 = 20
    ShiftF9 = 21
    ShiftF10 = 94
    ShiftF11 = 23
    ShiftF12 = 24

SapVKeyNames = {
    "ctrl+a": SapVKey.CtrlA,
    "ctrl+b": SapVKey.CtrlB,
    "ctrl+c": SapVKey.CtrlC,
    "ctrl+d": SapVKey.CtrlD,
    "ctrl+e": SapVKey.CtrlE,
    "ctrl+f": SapVKey.CtrlF,
    "ctrl+f1": SapVKey.CtrlF1,
    "ctrl+f2": SapVKey.CtrlF2,
    "ctrl+f3": SapVKey.CtrlF3,
    "ctrl+f4": SapVKey.CtrlF4,
    "ctrl+f5": SapVKey.CtrlF5,
    "ctrl+f6": SapVKey.CtrlF6,
    "ctrl+f7": SapVKey.CtrlF7,
    "ctrl+f8": SapVKey.CtrlF8,
    "ctrl+f9": SapVKey.CtrlF9,
    "ctrl+f10": SapVKey.CtrlF10,
    "ctrl+f11": SapVKey.CtrlF11,
    "ctrl+f12": SapVKey.CtrlF12,
    "ctrl+G": SapVKey.CtrlG,
    "ctrl+ins": SapVKey.CtrlIns,
    "ctrl+k": SapVKey.CtrlK,
    "ctrl+n": SapVKey.CtrlN,
    "ctrl+o": SapVKey.CtrlO,
    "ctrl+p": SapVKey.CtrlP,
    "ctrl+r": SapVKey.CtrlR,
    "ctrl+shift+0": SapVKey.CtrlShift0,
    "ctrl+shift+f1": SapVKey.CtrlShiftF1,
    "ctrl+shift+f2": SapVKey.CtrlShiftF2,
    "ctrl+shift+f3": SapVKey.CtrlShiftF3,
    "ctrl+shift+f4": SapVKey.CtrlShiftF4,
    "ctrl+shift+f5": SapVKey.CtrlShiftF5,
    "ctrl+shift+f6": SapVKey.CtrlShiftF6,
    "ctrl+shift+f8": SapVKey.CtrlShiftF8,
    "ctrl+shift+f9": SapVKey.CtrlShiftF9,
    "ctrl+shift+f10": SapVKey.CtrlShiftF10,
    "ctrl+shift+f11": SapVKey.CtrlShiftF11,
    "ctrl+shift+f12": SapVKey.CtrlShiftF12,
    "ctrl+t": SapVKey.CtrlT,
    "ctrl+v": SapVKey.CtrlV,
    "ctrl+x": SapVKey.CtrlX,
    "ctrl+y": SapVKey.CtrlY,
    "enter": SapVKey.Enter,
    "f1": SapVKey.F1,
    "f2": SapVKey.F2,
    "f3": SapVKey.F3,
    "f4": SapVKey.F4,
    "f5": SapVKey.F5,
    "f6": SapVKey.F6,
    "f7": SapVKey.F7,
    "f8": SapVKey.F8,
    "f9": SapVKey.F9,
    "f10": SapVKey.F10,
    "f11": SapVKey.F11,
    "f12": SapVKey.F12,
    "help": SapVKey.Help,
    "shift+f2": SapVKey.ShiftF2,
    "shift+f3": SapVKey.ShiftF3,
    "shift+f4": SapVKey.ShiftF4,
    "shift+f5": SapVKey.ShiftF5,
    "shift+f6": SapVKey.ShiftF6,
    "shift+f7": SapVKey.ShiftF7,
    "shift+f8": SapVKey.ShiftF8,
    "shift+f9": SapVKey.ShiftF9,
    "shift+f10": SapVKey.ShiftF10,
    "shift+f11": SapVKey.ShiftF11,
    "shift+f12": SapVKey.ShiftF12,
}

class SAPGui:

    def __init__(self, path: str=r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"):
        self.sap_path = path
        self.sapVKey = SapVKey

        if isinstance(client.GetObject('SAPGUI'), client.CDispatch):
            self.SapGuiAuto = client.GetObject('SAPGUI')

    def open_connection(self, conn: str, start_app: bool=False, timeout: Union[int, float]=30):

        if start_app:
            self.start_app()

        self.application = self._get_application(timeout)
        self.connection = self.application.OpenConnection(conn)
        self.session = self.connection.Children(0)
        self.wnd = self.session.findById("wnd[0]")
        return self

    def _get_application(self, timeout):
        for _ in TimeoutGenerator(timeout):
            try:
                return self.SapGuiAuto.GetScriptingEngine
            except Exception as e:
                message_error = e
        raise TimeoutError(message_error)
        
        
    def start_app(self):
        if not os.path.exists(self.sap_path):
            raise SAPNotFoundError(self.sap_path)
        subprocess.Popen(self.sap_path)


    def login(self, user:str, password:str, user_id:str="wnd[0]/usr/txtRSYST-BNAME", pass_id:str="wnd[0]/usr/pwdRSYST-BCODE"):
        self.session.findById(user_id).text = user
        self.session.findById(pass_id).text = password
        self.wnd.sendVKey(0)
        self.wnd.sendVKey(0)

    def get_object(self, id_object):
        return SAPObject(self.session.findById(id_object))

    def sendKey(self, key:str, vkey:int=None):
        if vkey is  None:
            vkey = SapVKeyNames[key]
        
        self.wnd.sendVKey(vkey)
        

class SAPObject:
    def __init__(self, sap_object):
        self.api = sap_object
        self.__create_properties(sap_object)

    def __create_properties(self, object):
        PROPERTIES = {
            "caretPosition": object.caretPosition,
            "close": object.close,
            "clickCurrentCells": object.clickCurrentCells,
            "contextMenu": object.contextMenu,
            "createSession": object.createSession,
            "currentCellColumn": object.currentCellColumn,
            "getcellvalue": object.getcellvalue,
            "Height": object.Height, 
            "Highlighted": object.Highlighted,
            "key": object.key,
            "maximize": object.maximize,
            "Name": object.Name,
            "press": object.press,
            "pressButton": object.pressButton,
            "pressContextButton": object.pressContextButton,
            "pressToolbarButton": object.pressToolbarButton,
            "pressToolbarContextButton": object.pressToolbarContextButton,
            "Required": object.Required,
            "sendVKey": object.sendVKey,
            "select": object.select,
            "selectColumn": object.selectColumn,
            "selectContextMenuItem": object.selectContextMenuItem,
            "selectNode": object.selectNode,
            "selected": object.selected,
            "selectedNode": object.selectedNode,
            "selectedRows": object.selectedRows,
            "setCurrentCell": object.setCurrentCell,
            "setFocus": object.setFocus,
            "Text": object.Text,
            "text": object.text,
            "verticalScrollbar": object.verticalScrollbar,
            "Width": object.Width,
            }
        self.__dict__.update(PROPERTIES)
        

    def get_cell(self, row: int, column: str):
        return self.getcellvalue(row, column)

    def click_cell(self, row:str, column:str)->None:
        self.currentCellColumn = column
        self.selectedRows = row
        self.clickCurrentCells()

class SAPNotFoundError(Exception):

    def __init__(self, path, message="SAP application not found."):
        self.path = path
        self.message = message

    def __str__(self):
        return self.message + f" Check if SAP is installed in this path: {self.path}"


class TimeoutGenerator:
    
    def __init__(self, max=None):
        self.max = max
    
    def __iter__(self):
        self.start = time.perf_counter()
        return self

    def __next__(self):
        delta = time.perf_counter() - self.start
        if not self.max or delta <= self.max:
            return delta
        raise StopIteration

if __name__ == "__main__":
        
    sap_object = SAPGui()
    sap_object.open_connection("test", timeout=10, start_app=True)
    sap_object.login("user", "12345")