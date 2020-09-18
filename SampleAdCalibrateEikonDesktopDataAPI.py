import pythoncom
import win32com.client
import datetime
import time
from enum import Enum
import sys

class ConnectionToEikonEnum(Enum):
        Error_InitializeFail          =2          # from enum EEikonDataAPIInitializeResult
        Error_Reinitialize            =1          # from enum EEikonDataAPIInitializeResult
        Succeed                       =0          # from enum EEikonDataAPIInitializeResult
        Connected                     =1          # from enum EEikonStatus
        Disconnected                  =0          # from enum EEikonStatus
        Disconnected_No_license_agreement_sign_off=8          # from enum EEikonStatus
        LocalMode                     =2          # from enum EEikonStatus
        Offline                       =4          # from enum EEikonStatus

#Calls to methods in AdfinX Analytics library can only be made while our application is connected to Eikon
#This variable is used to tell whether our Python application is connected to Eikon and therefore can make calls to AdfinX Analytics
connectedToEikon = False

class EikonDesktopDataAPI_EventHandler:
        def OnStatusChanged(self, EStatus):
                if EStatus == ConnectionToEikonEnum.Connected.value or EStatus == ConnectionToEikonEnum.LocalMode.value:
                        print("EikonDesktopDataAPI is connected in regular or local mode")
                        global connectedToEikon
                        connectedToEikon = True
                elif EStatus == ConnectionToEikonEnum.Disconnected.value:
                        print("EikonDesktopDataAPI is disconnected or not initialized")
                elif EStatus == ConnectionToEikonEnum.Disconnected_No_license_agreement_sign_off.value:
                        print("EikonDesktopDataAPI is disconnected because the user did not accept license agreement")
                elif EStatus == ConnectionToEikonEnum.Offline.value:
                        print("EikonDesktopDataAPI has lost connection to the platform due to network or platform issue")

#This creates an instance of EikonDesktopDataAPI object used to manage the connection between our Python application and Eikon
connectionToEikon = win32com.client.DispatchWithEvents("EikonDesktopDataAPILib.EikonDesktopDataAPI", EikonDesktopDataAPI_EventHandler)
print("Connecting to Eikon...")

if not connectedToEikon:
        retval = connectionToEikon.Initialize()
        if retval != ConnectionToEikonEnum.Succeed.value:
                print("Failed to initialize Eikon Desktop Data API")
                sys.exit()
else:
        print("Already connected to Eikon")        

while True:

        try:
                time.sleep(1) 
        except KeyboardInterrupt:
                print("KeyboardInterrupt")
                break

        #Windows message pump is required for COM objects to be able to raise events
        pythoncom.PumpWaitingMessages() 

        if connectedToEikon:
            #This creates an instance of AdxYieldCurveModule object from AdfinX Analytics library
            curveModule = connectionToEikon.CreateAdxYieldCurveModule()

            #Inputs required for AdCalibrate function
            rateStructure = "RM:HW ZCTYPE:RATE IM:CUBR"
            calcStructure = "CMT:FORM"

            inputArray = [[0 for x in range(9)] for x in range(2)]
            zeroCurve = [[0 for x in range(2)] for x in range(6)] 
    
            inputArray[0][0] = "S"
            inputArray[0][1] = pythoncom.MakeTime(datetime.date(2017,11,20))
            inputArray[0][2] = pythoncom.MakeTime(datetime.date(2017,12,20))
            inputArray[0][3] = "1Y"
            inputArray[0][4] = 0
            inputArray[0][5] = pythoncom.MakeTime(datetime.date(2017,12,20))
            inputArray[0][6] = 0.1
            inputArray[0][7] = "CALL EXM:E"
            inputArray[0][8] = "EUR_AB6E"
    
            inputArray[1][0] = "S"
            inputArray[1][1] = pythoncom.MakeTime(datetime.date(2017,11,20))
            inputArray[1][2] = pythoncom.MakeTime(datetime.date(2019,11,22))
            inputArray[1][3] = "2Y"
            inputArray[1][4] = 0
            inputArray[1][5] = pythoncom.MakeTime(datetime.date(2019,11,22))
            inputArray[1][6] = 0.11
            inputArray[1][7] = "CALL EXM:E"
            inputArray[1][8] = "EUR_AB6E"

            zeroCurve[0][0] = pythoncom.MakeTime(datetime.date(2017,11,21))
            zeroCurve[0][1] = 0.02
            zeroCurve[1][0] = pythoncom.MakeTime(datetime.date(2017,11,22))
            zeroCurve[1][1] = 0.02
            zeroCurve[2][0] = pythoncom.MakeTime(datetime.date(2018,5,22))
            zeroCurve[2][1] = 0.021
            zeroCurve[3][0] = pythoncom.MakeTime(datetime.date(2019,11,22))
            zeroCurve[3][1] = 0.022
            zeroCurve[4][0] = pythoncom.MakeTime(datetime.date(2027,11,23))
            zeroCurve[4][1] = 0.025
            zeroCurve[5][0] = pythoncom.MakeTime(datetime.date(2047,11,22))
            zeroCurve[5][1] = 0.035

            #Now AdCalibrate function can be called with above inputs
            try:
                calibratedRateArray = curveModule.AdCalibrate(inputArray,zeroCurve,[],rateStructure,calcStructure,"")
                print ("Calibrated rate array:")
                print (calibratedRateArray)
            except pythoncom.com_error as e:
                print (str(e))

            break
