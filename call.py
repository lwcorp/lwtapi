import sys
import win32com.client,time
tapi = win32com.client.Dispatch("TAPI.TAPI")
tapi.Initialize()
# for item in tapi.Addresses: print(item.AddressName)
modemAll = False
modem = "USB"

if len(sys.argv) > 1:
    sNumber = sys.argv[1]
else:
    exit("Send a phone number to the program")

found = False
modemList = []
for item in tapi.Addresses:
   if modemAll or modem in item.AddressName:
     found = True
     gobjCall = item.CreateCall(sNumber, 1, 0x8)
     gobjCall.Connect (False)
     if not modemAll:
        break
   else:
     modemList.append(item.AddressName)
if not found:
   modemList = str(modemList).strip('[]').replace(',', '\n')
   print(f"No modem matched {modem} in\n{modemList}")