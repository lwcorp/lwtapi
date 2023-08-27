import pythoncom, time
import win32com.client
cls="TAPI.TAPI"
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

def print_check_error(str, obj, code):
   try:
    output = obj.CallInfoString(code)
    if bool(output):
       if print_console:
          print(f"{str}: {output}")
       else:
          update_text(f"{str}:\n{output}")
   except:
     pass

def update_text(str):
    label.config(text=str)

tapi=win32com.client.Dispatch(cls)
tapi.Initialize() # must run after Dispatch and before TapiEvents
events=TapiEvents(tapi)
tapi.EventFilter = 0x1FFFF

for addr in tapi.Addresses: 
    try:
        tapi.RegisterCallNotifications(addr,True,True,8,0)
    except:
        pass

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