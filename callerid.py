import pythoncom, time
import win32com.client
cls="TAPI.TAPI"
modemAll = True
modem = "USB"
# need for gen_py
ti = win32com.client.Dispatch(cls)._oleobj_.GetTypeInfo()
tlb, index = ti.GetContainingTypeLib()
tla = tlb.GetLibAttr()
win32com.client.gencache.EnsureModule(tla[0], tla[1], tla[3], tla[4], bValidateFile=0)

print_console = False

class TapiEvents(win32com.client.getevents(cls)):
    def OnEvent(self, ev1,ev2): 
        constants = win32com.client.constants
        if ev1 == constants.TE_CALLNOTIFICATION:
            call = win32com.client.Dispatch(ev2).Call
            try:
               print_check_error(f"CIS_CALLEDIDNAME", call, constants.CIS_CALLEDIDNAME)
               print_check_error(f"CIS_CALLERIDNUMBER", call, constants.CIS_CALLERIDNUMBER)
               print_check_error(f"CIS_CALLEDIDNUMBER", call, constants.CIS_CALLEDIDNUMBER)
               print_check_error(f"CIS_CONNECTEDIDNAME", call, constants.CIS_CONNECTEDIDNAME)
               print_check_error(f"CIS_CONNECTEDIDNUMBER", call, constants.CIS_CONNECTEDIDNUMBER)
               print_check_error(f"CIS_REDIRECTIONIDNAME", call, constants.CIS_REDIRECTIONIDNAME)
               print_check_error(f"CIS_REDIRECTIONIDNUMBER", call, constants.CIS_REDIRECTIONIDNUMBER)
               print_check_error(f"CIS_REDIRECTINGIDNAME", call, constants.CIS_REDIRECTINGIDNAME)
               print_check_error(f"CIS_REDIRECTINGIDNUMBER", call, constants.CIS_REDIRECTINGIDNUMBER)
               print_check_error(f"CIS_CALLEDPARTYFRIENDLYNAME", call, constants.CIS_CALLEDPARTYFRIENDLYNAME)
               print_check_error(f"CIS_COMMENT", call, constants.CIS_COMMENT)
               print_check_error(f"CIS_DISPLAYABLEADDRESS", call, constants.CIS_DISPLAYABLEADDRESS)
               print_check_error(f"CIS_CALLINGPARTYID", call, constants.CIS_CALLINGPARTYID)
            except:
               pass

def print_check_error(name, obj, code):
    descriptions = {
        "CIS_CALLERIDNAME": "The name of the caller",
        "CIS_CALLERIDNUMBER": "The number of the caller",
        "CIS_CALLEDIDNAME": "The name of the called location",
        "CIS_CALLEDIDNUMBER": "The number of the called location",
        "CIS_CONNECTEDIDNAME": "The name of the connected location",
        "CIS_CONNECTEDIDNUMBER": "The number of the connected location",
        "CIS_REDIRECTIONIDNAME": "The name of the location to which a call has been redirected",
        "CIS_REDIRECTIONIDNUMBER": "The number of the location to which a call has been redirected",
        "CIS_REDIRECTINGIDNAME": "The name of the location that redirected the call",
        "CIS_REDIRECTINGIDNUMBER": "The number of the location that redirected the call",
        "CIS_CALLEDPARTYFRIENDLYNAME": "The called party friendly name",
        "CIS_COMMENT": "Comment about the call",
        "CIS_DISPLAYABLEADDRESS": "Displayable version of the address",
        "CIS_CALLINGPARTYID": "Identifier of the calling party"
    }

    try:
        output = obj.CallInfoString(code)
        if bool(output):  # Only display if there's a value
            description = descriptions.get(name, name)  # Get description, fallback to the key
            if print_console:
                print(f"{description}: {output}")
            else:
                update_text(f"{description}:\n{output}")
    except:
        pass

def update_text(str):
    label.config(text=str)

tapi=win32com.client.Dispatch(cls)
tapi.Initialize() # must run after Dispatch and before TapiEvents
events=TapiEvents(tapi)
tapi.EventFilter = 0x1FFFF

found = False
modemList = []
for addr in tapi.Addresses:
    try:
        if modemAll or modem in addr.AddressName:
            found = True
            tapi.RegisterCallNotifications(addr, True, True, 8, 0)
            if not modemAll:
                break
        else:
            modemList.append(addr.AddressName)
    except:
        pass
if not found:
    modemList = str(modemList).strip('[]').replace(',', '\n')
    print(f"No modem matched '{modem}' in:\n{modemList}")

if not print_console:
  try:
     import tkinter as tk
     r = tk.Tk()
     r.geometry(f"500x300+{int((r.winfo_screenwidth() / 2.5) - (r.winfo_reqwidth() / 2))}+{int((r.winfo_screenheight() / 2) - (r.winfo_reqheight() / 2))}")
     r.title("Incoming calls")
     label = tk.Label(r, text="Waiting for calls")
     label.pack()
     r.mainloop()
  except:
     print_console = True
     pass

if print_console:
  try:
     while True:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.01)  # Don't use up all our CPU checking constantly
  except KeyboardInterrupt:
     pass
