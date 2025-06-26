import tkinter as tk
import numpy as np
import pandas as pd
from tkinter import filedialog, messagebox, ttk
from tkinter import *
from tkinter.ttk import *
import re
import os

# Get the directory of the current Python script
script_directory = os.path.dirname(os.path.abspath(__file__))

# Set the current working directory to the directory of the script
os.chdir(script_directory)

def createDF():
    global df
    df = pd.DataFrame(columns=['Code', 'Description', 'Quantity', 'Price', 'Discount', 'Discount Price', 'Total'])

def Refresh():
    global df
    Emptentry = pd.DataFrame({"Code": [" "],
                            "Description": [" "],
                            "Quantity": [" "],
                            "Price": [" "],
                            "Discount": [" "],
                            "Discount Price": [" "],
                            "Total": [" "]})
    df = pd.concat([df, Emptentry])
    
    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text = column)
        
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tv1.insert("", "end", values=row)
    
    tv1.column("Code", width = 80)
    tv1.column("Description", width = 350)
    tv1.column("Quantity", width = 50, anchor=tk.CENTER)
    tv1.column("Price", width = 65, anchor=tk.CENTER)
    tv1.column("Discount", width = 40, anchor=tk.CENTER)
    tv1.column("Discount Price", width = 65, anchor=tk.CENTER)
    tv1.column("Total", width = 80, anchor=tk.CENTER)
    return None

def File_dialog():
    filename = filedialog.askopenfilename(initialdir="/", 
                                        title="Select A File", 
                                        filetypes=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file["text"] = filename
    return None
    
def Load_excel_data():
    File_dialog()
    MemberList()

    file_path = label_file["text"]
    global pricedf
    try:
        excel_filename = r"{}".format(file_path)
        pricedf = pd.read_excel(excel_filename)
        pricedf = pricedf.iloc[2:, 0:3]
    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid.")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    label_file.config(text = "Load successful")
    
    print(pricedf.head())
    
def getCustomerList():
    global CIDList
    CIDList = Customerdf.iloc[:,0]
    global CNameList
    CNameList = Customerdf.iloc[:,1]

    global CList
    CList = []

    for i in range(0, len(CIDList)):
        CID = str(CIDList[i])
        CName = str(CNameList[i])
        
        CString = CID + " - " + CName
        
        CList.append(CString)
    
def Load_Customer_excel_data():
    File_dialog()

    file_path = label_file["text"]
    global Customerdf
    try:
        excel_filename = r"{}".format(file_path)
        Customerdf = pd.read_excel(excel_filename)
        Customerdf = Customerdf.iloc[:, [0, 12]]
        getCustomerList()
        updateListBox(CList)
    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid.")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    
    label_file.config(text = "Load successful")
    print(Customerdf.head())
      
def Save_Excel():
    global df
    file = filedialog.asksaveasfilename(defaultextension = ".xlsx")
    df.to_excel(str(file))
    label_file.config(text = "File saved")
    
def clear_data():
    tv1.delete(*tv1.get_children())
    pass

def MemberList():
    global df
    global Costdf
    global RafterList
    global Rafterdf
    global R138RafterList
    global R138Rafterdf
    global R162Rafterdf
    global R162RafterList
    

    Costdf = pd.read_excel("Carport Member Rates.xlsx")

    Rafterdf = Costdf.loc[:, ['Rafter Code', 'Rafter Description', 'Rafter Weight [kg/m]']]
    Rafterdf = Rafterdf.dropna()
    R162Rafterdf = Rafterdf
    RafterList = Rafterdf.iloc[:,0].tolist()
    R162RafterList = RafterList

    R138Rafterdf = Costdf.loc[:, ['R138 Rafter Code', 'R138 Rafter Description', 'R138 Rafter Weight [kg/m]']]
    R138Rafterdf = R138Rafterdf.dropna()
    R138RafterList = R138Rafterdf.iloc[:,0].tolist()
    
    global SHS100df
    global SHS100CodeList
    SHS100df = Costdf.loc[:, ['100x100 SHS Code', '100x100 SHS Description', '100x100 SHS Weight [kg/m]']]
    SHS100df = SHS100df.dropna()
    SHS100CodeList = SHS100df.iloc[:, 0].tolist()
    
    global SHS76df
    global SHS76CodeList
    SHS76df = Costdf.loc[:, ['76x76 SHS Code', '76x76 SHS Description', '76x76 SHS Weight [kg/m]']]
    SHS76df = SHS76df.dropna()
    SHS76CodeList = SHS76df.iloc[:,0].tolist()
    
    global weightpm100
    weightpm100 = float(SHS100df.iloc[0,2])
    
    global weightpm76
    weightpm76 = float(SHS76df.iloc[0,2])
     
def rafter_choice_selected(event):
    # Get the selected value from RafterChoice
    selected_rafter_choice = MountVar.get()
    
    # Update the options in RafterLength based on the selected value
    RaftVar.set("")  # Clear the current selection
    RafterChoiceOp['menu'].delete(0, 'end')  # Clear previous options
    
    for option in rafter_length_options[selected_rafter_choice]:
        RafterChoiceOp['menu'].add_command(label=option, command=tk._setit(RaftVar, option))
    
def getMount():
    global Mount
    global MountS
    
    MountS = MountVar.get()
    
    if (MountS == 'R138 Rafter'):
        Mount = 0
    elif (MountS == 'R162 Rafter'):
        Mount = 1

def round_up(number):
    """
    Rounds a number up to the nearest integer.

    Args:
        number (float): The number to round up.

    Returns:
        int: The rounded up integer.
    """
    return math.ceil(number)

def extract_percentage_value(percentage_str):
    """
    Extracts the numeric value from a percentage string and returns it as an integer.

    Args:
        percentage_str (str): A string containing a percentage (e.g., "10%").

    Returns:
        int: The numeric value without the percentage sign.
    """
    return int(percentage_str.strip().replace('%', ''))

def getSmalls():
    global SmallSmalls
    global ConSmalls
    global SuppSmalls
    
    SmallSmalls = 1 + extract_percentage_value(SSmallsVar.get())/100
    ConSmalls = 1 + extract_percentage_value(ConSmallsVar.get())/100
    SuppSmalls = 1 + extract_percentage_value(SuppSmallsVar.get())/100

def getInputs():

    createDF()
    
    #get Numerical Inputs
    global sysnum
    global pHor
    global pVer
    global pWidth
    global pLength
    global GroundClearance
    global pNum
    global angle
    global PanelO

    sysnum = int(TableNumberE.get())
    PanelO = str(OrientationVar.get())
    pHor = int(HorPanelE.get())
    pVer = int(VertPanelE.get())
    pWidth = float(PanelWidthE.get()) 
    pLength = float(PanelLengthE.get())
    angle = np.deg2rad(float(AngleE.get()))
    GroundClearance = float(GroundClearanceE.get())

    #Vert Panels
    global pVert
    pVert = int(VertPanelE.get())
    pNum = pVert*pHor

    #get ROH
    global ROH
    ROHS = var.get()
    
    if(ROHS == '600mm'):
        ROH = 600
    elif(ROHS == '800mm'):
        ROH = 800
    selection = str(ROH) 

    #get Discount 
    global discount
    discount = 1 - float(DiscountE.get())/100

    #Type of Rafter
    global RafterType
    RafterType = MountVar.get()

    #Single or double access
    global SorD
    SorD = SDVar.get()
        
    #Number of Knee Braces
    global KneeBraces
    KneeBraces = int(KBVar.get())
        
    #Rafter Splice
    global RaftSplice
    RaftSplice = RaftSVar.get()

    #Front Rafter Overhang
    global FROHang
    FROHang = int(FRaftOvE.get())
    
    #Rear Rafter Overhang
    global RROHang
    RROHang = int(RRaftOvE.get())
    
    #Steel Rate
    global SteelRate
    SteelRate = float(SteelRateE.get())
    
    getSmalls()
    
def getRaftChoice():
    global RafterLChosen
    global Rafterdf
    global RafterDescr
    global RafterCode
    
    RafterLChosenStr = str(RaftVar.get())
    
    RafterList = Rafterdf['Rafter Description'].tolist()
    
    index = 0
    
    for i in range(2,len(RafterList)):
            if (RafterList[i] == RafterLChosenStr):
                index = i
    
    RafterDescr = RafterList[index]
    
    RafterCode = Rafterdf.iloc[index, 0]
    
    #Calculating the chosen Rafter Length
    global TotalRafterL
    Rafter1L = int(get_last_set_of_numbers(RafterCode))
    
    if(RaftSplice == 'Yes'):
        Rafter2Code = RaftVar2.get()
        Rafter2L = int(get_last_set_of_numbers(Rafter2Code))
        TotalRafterL = Rafter1L + Rafter2L
    else:
        TotalRafterL = Rafter1L
    
    RafterLChosen = TotalRafterL

def SHSprice(length, weightpm):
    
    SHSPrice = ((weightpm * length/1000 * SteelRate * 1.25)/0.6)*(1 - discountp/100)
    
    return round(SHSPrice, 2)

def RailCalc():
    global pVert
    global RailMult
    global sysnum
    global pHor
    global pWidth
    global pLength
    global pNum
    global df

    getInputs()
    getMount()
    #getMount()
    
    pNum = pVert*pHor
    
    if (PanelO == 'Landscape'):
        RailMult = pVert + 1
    elif (PanelO == 'Portrait'):
        RailMult = pVert * 2

def ClampCalc():

    RailCalc()

    global sysnum
    global pVert
    global pHor
    global RailMult
    global df
    global pNum


    pNum = pVert*pHor
    
    #End Clamp Calculations + Entry

    ECMult = (pHor//20 + 1)
    
    if (PanelO == 'Landscape'):
        EndClamps = pHor * 2 * 2 * sysnum
        
        InterClamps = (RailMult - 2)*pHor*2*sysnum
    elif (PanelO == 'Portrait'):
        EndClamps = (2*RailMult)*sysnum*ECMult
        
        #InterClamp Calculations 
        InterClamps = ((RailMult*(pHor - 1)))*sysnum
        

    AddEntry("LM-EC35-RNW", EndClamps, 0)
    AddEntry("LM-IC35-GP1-RNW", InterClamps, 0)
    
def split_value(value):
    # Calculate the larger part
    larger_part = round(value / 1.6)  # 100% + 30% + 30% = 160% => value / 1.6
    # Calculate the smaller parts
    smaller_part = round(larger_part * 0.3)
    return larger_part, smaller_part, smaller_part

def getPurlins():

    ClampCalc()

    global sysnum
    global pHor
    global pWidth
    global pLength
    global RafterLChosen
    global SupportLegs
    global pVert
    global df
    global Mount
    global RailMult
    global MountS
    global PurlinLMin
    
    if (PanelO == 'Landscape'):
        CalcPurlinL = pHor*pLength + 20*(pHor - 1)
    elif(PanelO == 'Portrait'):
        CalcPurlinL = pHor*pWidth + 20*(pHor - 1) + 100
    
    CalcPurlLabel.config(text = "Required Purlin Length (mm): " +str(CalcPurlinL))
    
    maxnum = 39
    curnum = 0

    PurlinL = CalcPurlinL + 300
    global PurlinLMin
    #PurlinLMin = PurlinL
    #PurlinLCalc = PurlinL

    L6200num = CalcPurlinL//6200 - 1
    L5250num = CalcPurlinL//5250 - 1
    Purlin6200 = L6200num * 6200
    ShortCalcPurlin = CalcPurlinL - Purlin6200
    #ShortCalcPurlin = CalcPurlinL - Purlin5250
    PurlinLMint = ShortCalcPurlin + 5000
    
    
    L1 = 5250
    L2 = 4400
    L3 = 3300

    L1num = 0
    L2num = 0
    L3num = 0
    
    L1numt = 0
    

    for i in range(0,maxnum):
        for j in range(0,maxnum):
            for k in range(0,maxnum):
                
                PurlinLCalc = i*L1 + j*L2 + k*L3
                
                if (PurlinLCalc<PurlinLMint and PurlinLCalc>ShortCalcPurlin):
                    
                    L1numt = (i)*RailMult*sysnum 
                    L2num = j*RailMult*sysnum 
                    L3num = k*RailMult*sysnum 
                    PurlinLMint = PurlinLCalc
                    
    #PurlinLMin = Purlin6200 + PurlinLMint
    PurlinLMin  = PurlinL
    L0num = L6200num*RailMult*sysnum
    
    if(L0num>0):
        AddEntry("LM-R112-W-6200", L0num, 0)
    
    L1num =  L1numt
    
    if(L1num>0):
        AddEntry("LM-R112-W-5250", L1num, 0)
    
    if(L2num>0):
        AddEntry("LM-R112-W-4400", L2num, 0)
    
    if(L3num>0):
        AddEntry("LM-R112-W-3300", L3num, 0)
    
    PurlinSplicersNum = L0num + L1num + L2num + L3num - (RailMult * sysnum)
    AddEntry("LM-RS-I-R112-W-300", PurlinSplicersNum, 0)
    
        
    StitchingScrews = 18*(PurlinSplicersNum)
    AddEntry("FS-S-22X6-C4", StitchingScrews, 0)
   
    PurlinSuppString = "Supplied Purlin Length: " + str(PurlinLMin) + "mm"
    PurlinLabel.config(text = PurlinSuppString)
    
    # Support calculation for carports
    global bays5mcalc
    global bays2p5mcalc
    global POHang1
    global POHang2
    global POHangwarning
    
    POHang1 = 0
    POHang2 = 0
    
    POHangwarning = 0 #0 means there is no warning, 1 means there is a warning
    
    #calculating the base number of 5m bays that can fit in the required purlin length
    bays2p5mcalc = 0
    bays5mcalc = PurlinLMin//5000 - 1
    if(PurlinLMin <= 8600):
        bays5mcalc = 1
        bay, POHang1, POHang2 = split_value(PurlinLMin)
        remaininglength = 10500
    elif(PurlinLMin > 8600 and PurlinLMin < 10000):
        bays5mcalc = 1
        bays2p5mcalc = 1
        remaininglength = PurlinLMin - bays5mcalc*5000 - bays2p5mcalc*2500
        POHang1 = remaininglength - 7500
        POHang2 = 0
    else:
        #calculating the remaining length after subtracting the sure number of 5m bays
        remaininglength = PurlinLMin - bays5mcalc*5000
       
    
    
    
    # Decision tree to make the best use of the the remaining purlin length
    if (bays5mcalc >= 0): #This is to check if there is at least 1 x 5m bay for sure
        if (remaininglength <= 2000): #This means that the remaining length can only supply an adequate overhang for 2.5m bays
            if (bays2p5mcalc < 2):
                bays5mcalc = bays5mcalc - 1
                bays2p5mcalc = 2
                remaininglength = PurlinLMin - bays5mcalc*5000 - bays2p5mcalc*2500
                POHang1 = round (remaininglength/2)
                POHang2 = round (remaininglength/2)
                
        elif (remaininglength > 2000 and remaininglength <= 3600): #This means that the remaining length can be used as overhangs for 5m bays
            remaininglength = PurlinLMin - bays5mcalc*5000 - bays2p5mcalc*2500
            POHang1 = round (remaininglength/2)
            POHang2 = round (remaininglength/2)
                
        elif (remaininglength > 3600 and remaininglength <= 5200): #This means that a 2.5m bay can be added and the overhang split between a 5m bay and a 2.5m bay
            bays2p5mcalc = 1
            remaininglength = PurlinLMin - bays5mcalc*5000 - bays2p5mcalc*2500
            POHang1 = round(remaininglength - 0.36*5000)
            POHang2 = remaininglength - POHang1
        
        elif (remaininglength > 5200 and remaininglength <= 6100): #This means that a 2.5m bay can be added in the middle so the overhang is still for 2 x 5m bays
            if(bays5mcalc >= 2):#checks if there's at least 2 x 5m bays
                bays2p5mcalc = 1
                remaininglength = PurlinLMin - bays5mcalc*5000 - bays2p5mcalc*2500
                POHang1 = round (remaininglength/2)
                POHang2 = round (remaininglength/2)
            elif ((remaininglength - 5000) > 5000): #checks if 2 x 5m bays can fit with minimum overhangs
                bays5mcalc = 2
                remaininglength = PurlinLMin - bays5mcalc*5000 - bays2p5mcalc*2500
                POHang1 = round (remaininglength/2)
                POHang2 = round (remaininglength/2)
                POHangwarning = 1
            elif ((remaininglength - 5000) <= 5000 and (remaininglength - 5000) > 2500): #checks if a 2.5m bay can fit with overhang warning
                bays5mcalc = 1
                bays2p5mcalc = 1
                remaininglength = PurlinLMin - bays5mcalc*5000 - bays2p5mcalc*2500
                POHang1 = round (remaininglength/2)
                POHang2 = round (remaininglength/2)
                POHangwarning = 1
            elif ((remaininglength - 5000) <= 5000 and (remaininglength - 5000) <= 2500): #last resort but with purlin overhang warning
                bays2p5mcalc = 1
                bays2p5mcalc = 0
                remaininglength = PurlinLMin - bays5mcalc*5000 - bays2p5mcalc*2500
                POHang1 = round (remaininglength/2)
                POHang2 = round (remaininglength/2)
                POHangwarning = 1
        elif (remaininglength > 6100 and remaininglength <= 7000):
            if(bays2p5mcalc >= 1):#checks if there's at least 1 x 2.5m bays
                bays5mcalc = bays5mcalc + 1
                remaininglength = PurlinLMin - bays5mcalc*5000 - bays2p5mcalc*2500
                POHang1 = round (5000*0.36)
                POHang2 = round (remaininglength - POHang1)
            elif(bays2p5mcalc < 1):
                bays2p5mcalc = 2
                remaininglength = PurlinLMin - bays5mcalc*5000 - bays2p5mcalc*2500
                POHang1 = round (5000*0.36)
                POHang2 = round (remaininglength - POHang1)
        elif (remaininglength > 7000 and remaininglength <= 8600):
            bays5mcalc = bays5mcalc + 1 
            remaininglength = PurlinLMin - bays5mcalc*5000 - bays2p5mcalc*2500
            POHang1 = round (remaininglength/2)
            POHang2 = round (remaininglength/2)
        elif (remaininglength > 8600 and remaininglength <= 10000):
            bays5mcalc = bays5mcalc + 1
            bays2p5mcalc = 1
            remaininglength = PurlinLMin - bays5mcalc*5000 - bays2p5mcalc*2500
            POHang1 = round (5000*0.36)
            POHang2 = round (remaininglength - POHang1)
            
        
    global PurlinOHang
    PurlinOHang = POHang1 + POHang2

    #Previous attempt at support spacing calculation
    global SupportLegsC
    SupportLegsC = 1 + (bays5mcalc + bays2p5mcalc)
    

    global SupportLegs
    SupportLegs = SupportLegsC * sysnum

    #Adding Purlin to Rafter Connectors
    global PRC
    totalrails = RailMult
    PRC = ((totalrails * SupportLegsC)*4)*sysnum
    
    if(RafterType == 'R138 Rafter'):
        PRCCode = "LMK-PRC-38"
        AddEntry(PRCCode, PRC, 0)
        CAPM8x20 = PRC
        AddEntry("FS-CAP-M8X20", CAPM8x20, 0)
        FWM8 = 2 * CAPM8x20
        AddEntry("FS-FW-M8", FWM8, 0)
        SWM8 = CAPM8x20
        AddEntry("FS-SW-M8", SWM8, 0)
        SQNM8 = CAPM8x20
        AddEntry("FS-SQN-M8", SQNM8, 0)
    else:
        PRCCode = "LM-PRC"
        AddEntry(PRCCode, PRC, 0)
    
    
    #Display the required Rafter Length
    if (PanelO == 'Landscape'):
        CalcRafterL = pVert * pWidth + ((pVert - 1)*20) + 200
    elif (PanelO == 'Portrait'):
        CalcRafterL = pVert * pLength + ((pVert - 1)*20) - ROH
        
    if (RaftSplice == 'Yes'):
        secondRafterEntry()
    else:
        RafterChoiceLabel.config(text = "Please select a standard Rafter Lengths in mm:")
    
    selection = "Calculated Rafter length: "+str(CalcRafterL)
    CalcRaftLabel.config(text = selection)
    
    SupportString = "5m bays: " + str(bays5mcalc) + " 2.5m bays: " + str(bays2p5mcalc)
    SupportSLabel.config(text = SupportString)

    PurlinSuppString = "Structure Length: " + str(PurlinLMin) + "mm"
    PurlinLabel.config(text = PurlinSuppString)
    
    SupportLegsStr = "Number of Support Legs per structure: " + str(SupportLegsC)
    SupportLegsLabel.config(text = SupportLegsStr)
    
    OHangString = "Purlin Overhang 1: "+str(POHang1)+"mm. \nPurlin Overhang 2: " + str(POHang2) + "mm"
    OHangLabel.config(text = OHangString)
    
    selection = "Calculated Rafter length: "+str(CalcRafterL)
    CalcRaftLabel.config(text = selection)
       
def get_last_set_of_numbers(input_string):
    # Find all sequences of digits in the input string
    matches = re.findall(r'\d+', input_string)
    
    # Return the last match if there are any, otherwise return None
    return matches[-1] if matches else None

def secondRafterEntry():
    RafterChoiceLabel.config(text = "Please select your rafters: ")
    
    global RaftVar2
    RaftVar2 = tk.StringVar()
    RaftStr2 = Rafterdf['Rafter Code'].tolist()
    RaftVar2.set(RaftStr2[0])
    RafterChoice2Op = tk.OptionMenu(InputFrame, RaftVar2, *RaftStr2)
    RafterChoice2Op.grid(row = 9, column = 3, padx = 5, pady = 5)
    
def RafterEntry(quantity):
    global df
    global RafterCode
        
    if (RaftSplice == 'Yes'):
        global Rafter2Code
        
        Rafter2Code = RaftVar2.get()
        RafterCode = RaftVar.get()
        
        r1description, r1price, r1discprice, r1total = getprice(RafterCode, quantity, 0)
        discountp = float(DiscountE.get())
        RaftEntry = pd.DataFrame({"Code": [RafterCode], 
                                "Description": [r1description],
                                "Quantity": [quantity], 
                                "Price": [r1price],
                                "Discount": [str(discountp) + "%"],
                                "Discount Price": [r1discprice],
                                "Total": [r1total]})
        df = pd.concat([df, RaftEntry])
        
        r2description, r2price, r2discprice, r2total = getprice(Rafter2Code, quantity, 0)
        discountp = float(DiscountE.get())
        RaftEntry = pd.DataFrame({"Code": [Rafter2Code], 
                                "Description": [r2description],
                                "Quantity": [quantity], 
                                "Price": [r2price],
                                "Discount": [str(discountp) + "%"],
                                "Discount Price": [r2discprice],
                                "Total": [r2total]})
        df = pd.concat([df, RaftEntry])
        
        raftersplicenum = quantity
        AddEntry("LM-CP-SB-MILL-L", raftersplicenum, 1000)
        FSHBM16x150 = quantity * 6
        AddEntry("FS-HB-M16X150", FSHBM16x150, 0)
        FSFWM16 = FSHBM16x150 * 2
        AddEntry("FS-FW-M16", FSFWM16, 0)
        FSSWM16 = FSHBM16x150
        AddEntry("FS-SW-M16", FSSWM16, 0)
        FSNM16 = FSHBM16x150
        AddEntry("FS-N-M16", FSNM16, 0)
        
    elif (RaftSplice == 'No'):
        RafterCode = RaftVar.get()
        
        r1description, r1price, r1discprice, r1total = getprice(RafterCode, quantity, 0)
        discountp = float(DiscountE.get())
        RaftEntry = pd.DataFrame({"Code": [RafterCode], 
                                "Description": [r1description],
                                "Quantity": [quantity], 
                                "Price": [r1price],
                                "Discount": [str(discountp) + "%"],
                                "Discount Price": [r1discprice],
                                "Total": [r1total]})
        df = pd.concat([df, RaftEntry])
    
    #Calculating the chosen Rafter Length
    global TotalRafterL
    Rafter1L = int(get_last_set_of_numbers(RafterCode))
    
    if(RaftSplice == 'Yes'):
        Rafter2L = int(get_last_set_of_numbers(Rafter2Code))
        TotalRafterL = Rafter1L + Rafter2L
    else:
        TotalRafterL = Rafter1L
    
def MountSupp():

    global sysnum
    global Mount
    global SupportLegs
    global pVert
    global RafterLChosen
    global RailMult
    global df
    global RafterLChosen
    global Rafterdf
    global RafterDescr
    global RafterCode
    
    RafterQuantity = SupportLegs
    
    RafterEntry(RafterQuantity)

    global FPOhang
    FPOhang = 500
        
    if(PanelO == 'Landscape'):
        FPOhang = 0
    
    #Calculates the length of the front support by taking the panel and rafter overhang into account
    global FrontSupport
    FrontSupport = round((FPOhang + FROHang)*np.sin(angle)) + GroundClearance 
    
    #Calculating the length of the rear support 
    global RearSupport
    DistBetwFandB = TotalRafterL - (FROHang + RROHang)
    RearSupport = FrontSupport + round(DistBetwFandB*np.sin(angle))
    
    global CentralSupport
    if (SorD == "Double-access"): #Check if it's double-access to calculate the central support length
        CentralSupport = round(FrontSupport + (RearSupport - FrontSupport)/2)  
        
    if (RafterType == "R162 Rafter"): #Standard 100x100 R162 Rafter Economax
        RLCMult = 2
        
        #Front Support Entry
        discountp = float(DiscountE.get())
        FSuppCode = SHS100CodeList[0]
        FSuppDescr = "ECONOMAX CARPORT FRONT COLUMN DIM: 100x100x3 LENGTH: "+ str(FrontSupport) +"mm PROFILE: SHS  FINNISH: HDG MATERIAL: S355JR"
        FSuppPrice = SHSprice(FrontSupport, weightpm100)
        FSuppDPrice = FSuppPrice * (1 - discountp/100)
        FSuppEntry = pd.DataFrame({"Code": [FSuppCode], 
                                "Description": [FSuppDescr],
                                "Quantity": [SupportLegs], 
                                "Price": [FSuppPrice],
                                "Discount": [str(discountp) + "%"],
                                "Discount Price": [FSuppDPrice],
                                "Total": [FSuppDPrice*sysnum]})
        df = pd.concat([df, FSuppEntry])
        
        if (SorD == "Double-access"): #Check if it's double-access to add the central support length
            RLCMult = 3
            CSuppCode = SHS100CodeList[1]
            CSuppDescr = "ECONOMAX CARPORT CENTRE COLUMN DIM: 100x100x3 LENGTH: "+ str(CentralSupport) +"mm PROFILE: SHS FINNISH: HDG MATERIAL: S355JR"
            CSuppPrice = SHSprice(CentralSupport, weightpm100)
            CSuppDPrice = CSuppPrice * (1 - discountp/100)
            CSuppEntry = pd.DataFrame({"Code": [CSuppCode], 
                                    "Description": [CSuppDescr],
                                    "Quantity": [SupportLegs], 
                                    "Price": [CSuppPrice],
                                    "Discount": [str(discountp) + "%"],
                                    "Discount Price": [CSuppDPrice],
                                    "Total": [CSuppDPrice*sysnum]})
            df = pd.concat([df, CSuppEntry])
        
        #Rear Support Entry
        RSuppCode = SHS100CodeList[2]
        RSuppDescr = "ECONOMAX CARPORT REAR COLUMN DIM: 100x100x3 LENGTH: "+ str(RearSupport) +"mm PROFILE: SHS  FINNISH: HDG MATERIAL: S355JR"
        RSuppPrice = SHSprice(RearSupport, weightpm100)
        RSuppDPrice = RSuppPrice * (1 - discountp/100)
        RSuppEntry = pd.DataFrame({"Code": [RSuppCode], 
                                "Description": [RSuppDescr],
                                "Quantity": [SupportLegs], 
                                "Price": [RSuppPrice],
                                "Discount": [str(discountp) + "%"],
                                "Discount Price": [RSuppDPrice],
                                "Total": [RSuppDPrice*sysnum]})
        df = pd.concat([df, RSuppEntry])    
        
        KneeBraceRLC = 0
        
        if(KneeBraces > 0):    
            KneeBracesqty = KneeBraces * SupportLegs
            AddEntry("LM-CP-SB-MILL-L", KneeBracesqty, 1500)
            
            FSHBM16x150 = KneeBracesqty * 1
            AddEntry("FS-HB-M16X150", FSHBM16x150, 0)
            FSFWM16 = FSHBM16x150 * 2
            AddEntry("FS-FW-M16", FSFWM16, 0)
            FSSWM16 = FSHBM16x150
            AddEntry("FS-SW-M16", FSSWM16, 0)
            FSNM16 = FSHBM16x150
            AddEntry("FS-N-M16", FSNM16, 0)
            
            KneeBraceRLC = KneeBracesqty
            
        #Adding the RLCs
        RLCqty = RLCMult * SupportLegs + KneeBraceRLC
        AddEntry("LM-CP-RLC-1", RLCqty, 0)
        
    elif(RafterType == "R138 Rafter"):
        RLCMult = 2 #How many RLCs needed for each upright member
        
        #Front Support Entry
        discountp = float(DiscountE.get())
        FSuppCode = SHS76CodeList[0]
        FSuppDescr = "ECONOMAX CARPORT FRONT COLUMN DIM: 76x76x3 LENGTH: "+ str(FrontSupport) +"mm PROFILE: SHS FINNISH: HDG MATERIAL: S355JR"
        FSuppPrice = SHSprice(FrontSupport, weightpm76)
        FSuppDPrice = FSuppPrice * (1 - discountp/100)
        FSuppEntry = pd.DataFrame({"Code": [FSuppCode], 
                                "Description": [FSuppDescr],
                                "Quantity": [SupportLegs], 
                                "Price": [FSuppPrice],
                                "Discount": [str(discountp) + "%"],
                                "Discount Price": [FSuppDPrice],
                                "Total": [FSuppDPrice*sysnum]})
        df = pd.concat([df, FSuppEntry])
        
        if (SorD == "Double-access"): #Check if it's double-access to add the central support length
            RLCMult = 3 #added 1 extra RLC for the centre support
            CSuppCode = SHS76CodeList[1]
            CSuppDescr = "ECONOMAX CARPORT CENTRE COLUMN DIM: 76x76x3 LENGTH: "+ str(CentralSupport) +"mm PROFILE: SHS FINNISH: HDG MATERIAL: S355JR"
            CSuppPrice = SHSprice(CentralSupport, weightpm76)
            CSuppDPrice = CSuppPrice * (1 - discountp/100)
            CSuppEntry = pd.DataFrame({"Code": [CSuppCode], 
                                    "Description": [CSuppDescr],
                                    "Quantity": [SupportLegs], 
                                    "Price": [CSuppPrice],
                                    "Discount": [str(discountp) + "%"],
                                    "Discount Price": [CSuppDPrice],
                                    "Total": [CSuppDPrice*sysnum]})
            df = pd.concat([df, CSuppEntry])
        
        #Rear Support Entry
        RSuppCode = SHS76CodeList[2]
        RSuppDescr = "ECONOMAX CARPORT REAR COLUMN DIM: 76x76x3 LENGTH: "+ str(RearSupport) +"mm PROFILE: SHS FINNISH: HDG MATERIAL: S355JR"
        RSuppPrice = SHSprice(RearSupport, weightpm76)
        RSuppDPrice = RSuppPrice * (1 - discountp/100)
        RSuppEntry = pd.DataFrame({"Code": [RSuppCode], 
                                "Description": [RSuppDescr],
                                "Quantity": [SupportLegs], 
                                "Price": [RSuppPrice],
                                "Discount": [str(discountp) + "%"],
                                "Discount Price": [RSuppDPrice],
                                "Total": [RSuppDPrice*sysnum]})
        df = pd.concat([df, RSuppEntry])    
        
        KneeBraceRLC = 0
        
        FSHBM12x110KB = 0
        
        if(KneeBraces > 0):    
            KneeBracesqty = KneeBraces * SupportLegs
            AddEntry("LM-SB-L", KneeBracesqty, 1500)
            
            FSHBM12x110KB = KneeBracesqty * 2
            
            KneeBraceRLC = KneeBracesqty
            
        #Adding the RLCs
        RLCqty = RLCMult * SupportLegs + KneeBraceRLC
        AddEntry("LM-CP-RLC-R138", RLCqty, 0)
        
        FSHBM12x110KB = RLCMult * SupportLegs + FSHBM12x110KB
        AddEntry("FS-HB-M12x110", FSHBM12x110KB, 0)
        FSFWM12 = FSHBM12x110KB * 2
        AddEntry("FS-FW-M12", FSFWM12, 0)
        FSSWM12 = FSHBM12x110KB
        AddEntry("FS-SW-M12", FSSWM12, 0)
        FSNM12 = FSHBM12x110KB
        AddEntry("FS-N-M12", FSNM12, 0)
        
        FSCAPM8x20 = RLCqty * 6
        AddEntry("FS-CAP-M8X20", FSCAPM8x20, 0)
        FSFWM8 = FSCAPM8x20 * 2
        AddEntry("FS-FW-M8", FSFWM8, 0)
        FSSWM8 = FSCAPM8x20
        AddEntry("FS-SW-M8", FSSWM8, 0)
        FSSQNM8 = FSCAPM8x20
        AddEntry("FS-SQN-M8", FSSQNM8, 0)
            
        
    #Calculating Purlin End Caps
    PurECs = RailMult * 2 * sysnum
    AddEntry("LM-R112-W-PEC", PurECs, 0)
        
    #Calculating Rafter End Caps
    RaftECs = 2 * SupportLegs
    if(RafterType == 'R138 Rafter'):
        RECCode = "LM-R138-REC"
    else:
        RECCode = "LM-CP-REC"
    AddEntry(RECCode, RaftECs, 0)
        
    #Adding Stitching Screws for End Caps
    ECStitchScr = (PurECs + RaftECs) * 2
    AddEntry("FS-S-22X6-C4", ECStitchScr, 0)
    
    FSTRM16x200 = RLCMult * 4 * SupportLegs
    AddEntry("FS-TR-M16x200-HDG", FSTRM16x200, 0)
    
    FSFWM16 = FSTRM16x200 * 2
    AddEntry("FS-FW-M16-HDG", FSFWM16, 0)
    
    FSNM16 = FSTRM16x200 * 2
    AddEntry("FS-N-M16-HDG", FSNM16, 0)
                
    # Cross Braces calculations:
    SupportSpaces = (SupportLegs/sysnum)-1
    NumberOfCrossSupport = (SupportSpaces//4 + 1)*sysnum # There should not be more than 4 spaces between supports. 5 is fine if only 2 are needed
    NumberOfCrossBraces = NumberOfCrossSupport*2

    #Adding the 6m support bars for the cross-bracing    
    
    AddEntry("LM-SB-6000", NumberOfCrossBraces, 0)
    
    if(RafterType == 'R138 Rafter'):
        FSHBM12x110KB = NumberOfCrossBraces
        AddEntry("FS-HB-M12x110", FSHBM12x110KB, 0)
        FSFWM12 = FSHBM12x110KB * 2
        AddEntry("FS-FW-M12", FSFWM12, 0)
        FSSWM12 = FSHBM12x110KB
        AddEntry("FS-SW-M12", FSSWM12, 0)
        FSNM12 = FSHBM12x110KB
        AddEntry("FS-N-M12", FSNM12, 0)

    elif(RafterType == 'R162 Rafter'):
        FSHBM16x150 = NumberOfCrossBraces
        AddEntry("FS-HB-M16X150", FSHBM16x150, 0)
        FSFWM16 = FSHBM16x150 * 2
        AddEntry("FS-FW-M16", FSFWM16, 0)
        FSSWM16 = FSHBM16x150
        AddEntry("FS-SW-M16", FSSWM16, 0)
        FSNM16 = FSHBM16x150
        AddEntry("FS-N-M16", FSNM16, 0)
        
    TTC90 = NumberOfCrossBraces
    AddEntry("LMK-TTC90-CBC", TTC90, 0)
    
    FP90 = NumberOfCrossBraces
    AddEntry("LM-FP-90", FP90, 0)
    
    FSTRM12x160 = FP90 * 2
    AddEntry("FS-TR-M12X160", FSTRM12x160, 0)
    FSFWM12 = FSTRM12x160
    AddEntry("FS-FW-M12", FSFWM12, 0)
    FSNM12 = FSTRM12x160
    AddEntry("FS-N-M12", FSNM12, 0)
    
    #adding IKA
    IKA70007 = (FSTRM16x200 + FSTRM12x160)//20 + 1
    AddEntry("IKA-70007", IKA70007, 0)
               
def replace_first_l_with_numbers(input_str, replacement_numbers):
    count = 0
    result = ''

    for char in input_str:
        if char == 'L':
            count += 1
            if count == 1:
                result += str(replacement_numbers)  # Replace 'L' with the desired numbers
            else:
                result += char
        else:
            result += char

    return result

def getprice(code, quantity, length):
    global pricedf
    global discountp
    global MarkUpe
    global MarkUp
    global RafterLChosen
    
    discountp = float(DiscountE.get())

    discount = 1 - discountp/100

    ref = pricedf.iloc[:,0]
    prices = pricedf.iloc[:,2]
    descriptions = pricedf.iloc[:, 1]
    
    string = code

    index = 0

    if (code == "LM-R110-4200"):
        price = 1214.36 
        RafterLChosen = 4200
        description = "Rafter 110x4200mm AL6005 T6 Mill"
    else:
        for i in range(2,len(ref)):
            if (ref[i] == string):
                index = i
                
        price = prices[index]
        description = descriptions[index]
    
    if(code == "LM-CP-SB-MILL-L"):
        pricet = (price*(length/1000)+40)
        description = "Carport Support Bar 118x" + str(length) + "mm AL6063 T6 Mill"
    elif(code == "LM-SB-L"):
        pricet = (float(price)*(length/1000)+17)
        description = "Support Bar 55x55x" + str(length) + "mm AL6063 T6 Mill"
    else:
        pricet = price

    price = round(float(pricet), 2)
    discprice = round(pricet*discount, 2)   
    totalprice = discprice*quantity
     
    
    return description, price, discprice, totalprice

def extract_length(s: str) -> int:
    
    if("Carport Support Bar" in s):
        match = re.search(r'\d+x(\d+)mm', s)

    else:
        match = re.search(r'\d+x\d+x(\d+)mm', s)
        
    if match:
        return int(match.group(1))
    raise ValueError("Length not found in string")    
    
def getStdSupportBarLength(Description):
    
    length = extract_length(Description)
    
    if("Carport Support Bar" in Description):
        SBLengthArray = [1075, 1100, 1180, 1265, 1300, 1360, 1370, 1410, 1450, 1500, 1620, 1690, 1710, 1740, 1780, 2130, 2250, 2280, 2300, 2320, 2440, 2480, 2510, 2530, 2580, 2700, 2750, 2930, 2980, 3390, 3400, 3425, 3440, 3497, 3580, 3610, 3670, 3680, 3715, 3730, 3820, 3905, 3940, 3945, 3990, 4000, 4020, 4050, 4180, 4270, 4320, 4360, 4420, 4580, 4760]
    else:
        SBLengthArray = [450, 530, 550, 590, 615, 1500, 1550, 1560, 1740, 1800, 1815, 1870, 1875, 1960, 2525, 2540, 2610, 2615, 2670, 2710, 2795, 2820, 2840, 2890, 3000, 3340, 5000, 6000]
    
    if (length <= SBLengthArray[0]):
        return SBLengthArray[0]
    
    else:
        for i in range(1, len(SBLengthArray)):
            if length <= SBLengthArray[i] and length > SBLengthArray[i-1]:
                return SBLengthArray[i]
                           
def AddK8Entry(code, quantity):
    global K8df
    
    NewEntry = pd.DataFrame({"Code": [code],
                             "Quantity": [str(quantity)]})
    K8df = pd.concat([K8df, NewEntry])

def ConvertToK8():
    global K8df
    
    K8df = pd.DataFrame(columns=['Code', 'Quantity'])
    
    K8Convertdf = pd.read_excel("Old and New Codes.xlsx")

    Oldcodedf = K8Convertdf.loc[:, ['Old Code']]
    Oldcodedf = Oldcodedf.dropna()
    OldcodeList = Oldcodedf.iloc[:,0].tolist()
    
    Newcodedf = K8Convertdf.loc[:, ['New Code']]
    Newcodedf = Newcodedf.dropna()
    NewcodeList = Newcodedf.iloc[:,0].tolist()
    
    QuoteCodesdf = df.loc[:, ['Code']]
    QuoteCodesdf = QuoteCodesdf.dropna()
    QuoteCodes = QuoteCodesdf.iloc[:,0].tolist()
    
    QuoteDescdf = df.loc[:, ['Description']]
    QuoteDescdf = QuoteDescdf.dropna()
    QuoteDescs = QuoteDescdf.iloc[:,0].tolist() 
    
    QuoteQuantitiesdf = df.loc[:, ['Quantity']]
    QuoteQuantitiesdf = QuoteQuantitiesdf.dropna()
    QuoteQuantities = QuoteQuantitiesdf.iloc[:,0].tolist()
    
    for i in range(0, len(QuoteCodes) - 1):
        for j in range(0, len(OldcodeList) - 1):
            
            if (QuoteCodes[i] == "LM-SB-L"):
                length = getStdSupportBarLength(QuoteDescs[i])
                StdSuppBarCode = "LM-SB-" + str(length)
                QuoteCodes[i] = StdSuppBarCode
                
            elif (QuoteCodes[i] == "LM-CP-SB-MILL-L"):
                length = getStdSupportBarLength(QuoteDescs[i])
                StdSuppBarCode = "LM-CP-SB-MILL-" + str(length)
                QuoteCodes[i] = StdSuppBarCode 
                     
            if (QuoteCodes[i] == OldcodeList[j]):
                
                AddK8Entry(NewcodeList[j], QuoteQuantities[i])

def LoadWeights():
    global Weightdf
    
    Weightdf = pd.read_excel("Inventory Volume & weight.xlsx")
    Weightdf = Weightdf.iloc[8:, 0:3].reset_index(drop=True)
    Weightdf.columns = Weightdf.iloc[0]
    Weightdf = Weightdf.iloc[1:, :].reset_index(drop=True)
    Weightdf = Weightdf.iloc[:, [0, 2]]
    Weightdf = Weightdf.dropna()
    
    global WeightCode
    global Weights
    WeightCode = Weightdf.iloc[:,0].tolist()
    Weights = Weightdf.iloc[:,1].tolist()

def getWeight(code, description, quantity):
    global WeightCode
    global Weights
    
    for i in range(0, len(WeightCode) - 1):
        if (code == WeightCode[i]):
            if (code == "LM-SB-L"):
                Length = extract_length(description)/1000
                weight = round(Weights[i] * Length)
            else:
                weight = Weights[i]
            break
        elif ("LM-GM-P-F" in code):
            weight = round(Weightpm * FrontSupport/1000)
        elif ("LM-GM-P-R" in code):
            weight = round(Weightpm * BackSupport/1000)
        elif (code == "LM-GM-CB-LC50"):
            weight = round(Bracewpm * BraceL/1000)
        else:
            weight = 0
    
    TotWeight = weight * int(quantity)
    TotWeight = round(float(TotWeight))
            
    return weight, TotWeight

def AddWeightEntry(weight, TotWeight):
    global quote_weight_df
    
    NewEntry = pd.DataFrame({"Unit Weight [g]": [float(weight)],
                             "Total Weight [g]": [float(TotWeight)]})
    quote_weight_df = pd.concat([quote_weight_df, NewEntry])

def CreateWeightDF():
    LoadWeights()
    
    global df
    global Weightdf
    
    # Create a new DataFrame with the same index as df
    global quote_weight_df
    quote_weight_df = pd.DataFrame(columns=['Unit Weight [g]', 'Total Weight [g]'])
    
    for i in range(1, len(df) - 1):
        code = df.iloc[i]['Code']
        description = df.iloc[i]['Description']
        quantity = df.iloc[i]['Quantity']
        
        weight, TotWeight = getWeight(code, description, quantity)
        
        # Add the weights to the new DataFrame
        AddWeightEntry(weight, TotWeight)
    
    #Adding total weight of the order    
    OrderWeight = (quote_weight_df['Total Weight [g]'].sum())/1000
    OrderWeight = round(float(OrderWeight), 3)
    OrderWeight = str(OrderWeight) + " kg"
    #AddWeightEntry("Total weight of the order", OrderWeight)
    NewEntry = pd.DataFrame({"Unit Weight [g]": ["Total weight of the order"],
                             "Total Weight [g]": [str(OrderWeight)]})
    quote_weight_df = pd.concat([NewEntry, quote_weight_df])
    
def CombineDataFrames(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    global df
    global K8df
    
    # Combine the two DataFrames
    df1 = df1.reset_index(drop=True)
    df2 = df2.reset_index(drop=True)
    
    return pd.concat([df1, df2], axis=1, ignore_index=False)


def AddEntry(code, quantity, length):
    global df
    global discountp
    
    description, price, discprice, total = getprice(code, quantity, length)
    
    NewEntry = pd.DataFrame({"Code": [code], 
                            "Description": [str(description)],
                            "Quantity": [quantity], 
                            "Price": [price],
                            "Discount": [str(discountp)+"%"],
                            "Discount Price": [discprice],
                            "Total": [total]})
    df = pd.concat([df, NewEntry])
       
def Calculations():
    debug = "Lol"
    debugLabel.config(text = debug)
    getPurlins()
    
def getDescription():
    global df
    global sysnum
    global angle
    global GroundClearance
    global pHor
    global pVert
    global pLength
    global pWidth
    global PurlinLMin
    global SupportLegsC
    
    description = "Table details: Table count: "+str(sysnum)+" Economax 4 post, "+str(RafterType)+", "+str(SorD)+", "+str(pVer)+"x"+str(pHor)+", "+str(PanelO)+", Support Count: "+str(SupportLegsC)
    description = description + ", Table length: "+str(PurlinLMin)+"mm, Purlin Runs: "+str(RailMult)+", for panel dimensions: "+str(pLength)+"x"+str(pWidth)+"mm, with "+str(bays5mcalc)+" 5m bays"
    description = description +" and "+str(bays2p5mcalc)+" 2.5m bays."
    
    total = round((df['Total'].sum()), 2)
    
    Descrentry = pd.DataFrame({"Code": ["DESCRIPTION"], 
                            "Description": [description],
                            "Quantity": [sysnum], 
                            "Price": [" "],
                            "Discount": [str(discountp)+"%"],
                            "Discount Price": [" "],
                            "Total": [total]})
    df = pd.concat([Descrentry, df])
    
def FinishCalc():
    
    getRaftChoice()
    MountSupp()
    getDescription()

    selection = "Success"
    debugLabel.config(text = selection)
    
def updateListBox(data):
    # clear list box
    ClientListBox.delete(0, END)
    
    # Add Clients to list box
    for item in data:
        ClientListBox.insert(END, item)
        
#Update entry box with listbox clicked
def fillout(e):
    #delete whatever is in the entry box
    CCodeE.delete(0, END)
    
    # Add clicked list item to entry box
    CCodeE.insert(0, ClientListBox.get(ACTIVE))
    
# Create function to check entry vs listbox
def check(e):
    # grab what was typed
    typed = CCodeE.get()
    
    if typed =='':
        data = CList
        updateListBox(data)
    else:
        data = []
        for item in CList:
            if typed.lower() in item.lower():
                data.append(item)
    
    updateListBox(data)
            
def ProjectInfo():
    # Toplevel object which will 
    # be treated as a new window
    global newWindow
    newWindow = Toplevel(root)
 
    # sets the title of the
    # Toplevel widget
    newWindow.title("Project Information Entry")
 
    # sets the geometry of toplevel
    newWindow.geometry("1380x750")
    
    #Customer List collect
    CustomerListFrame = tk.LabelFrame(newWindow, text = "Load the latest customer list")
    CustomerListFrame.grid(row = 0, column = 0, padx = 5, pady = 5)
    
    #Customer Information
    CCodeLabel = tk.Label(CustomerListFrame, text = "Customer Details:")
    CCodeLabel.grid(row = 1, column = 0, padx = 5, pady = 5)
    global CCodeE
    CCodeE = tk.Entry(CustomerListFrame, width = 75)
    CCodeE.grid(row = 1, column = 1, padx = 5, pady = 5)
    global ClientListBox
    ClientListBox = tk.Listbox(CustomerListFrame, width = 75)
    ClientListBox.grid(row = 2, column = 1, padx = 5, pady = 5)
    
    # Create a binding on the listbox on click
    ClientListBox.bind("<<ListboxSelect>>", fillout)
    
    # Create a binding on the entry box
    CCodeE.bind("<KeyRelease>", check)
    
    #Button to find customer file
    CustomerListB = tk.Button(CustomerListFrame, text = "Load Customer List", command = lambda: Load_Customer_excel_data())
    CustomerListB.grid(row = 0, column = 0, padx = 5, pady = 5)
    
    #Project Details Entries
    PDFrame  = tk.LabelFrame(newWindow, text = "Enter the project details")
    PDFrame.grid(row = 0, column = 1, padx = 5, pady = 5)
    
    #Date
    DateLabel = tk.Label(PDFrame, text = "Please enter today's date (YYYY/MM/DD):")
    DateLabel.grid(row = 0, column = 0, padx = 5, pady = 5)
    global DateE
    DateE = tk.Entry(PDFrame)
    DateE.grid(row = 0, column = 1, padx = 5, pady = 5)
    
    #Reference
    ReferenceLabel = tk.Label(PDFrame, text = "Enter Quote Reference:")
    ReferenceLabel.grid(row = 1, column = 0, padx = 5, pady = 5)
    global ReferenceE
    ReferenceE = tk.Entry(PDFrame, width = 75)
    ReferenceE.grid(row = 1, column = 1, padx = 5, pady = 5)
    
    #Message
    MessageLabel = tk.Label(PDFrame, text = "Enter Quote Message:")
    MessageLabel.grid(row = 2, column = 0)
    global MessageE
    MessageE = tk.Entry(PDFrame, width = 75)
    MessageE.grid(row = 2, column = 1, padx = 5, pady = 5)

    #Buttons
    ButtonFrame  = tk.LabelFrame(newWindow, text = "Capture Information")
    ButtonFrame.grid(row = 3, column = 0, padx = 5, pady = 5)
    
    PIButton = tk.Button(ButtonFrame, text = "Capture Project Information", command = lambda: getProjectInfo())
    PIButton.grid(row = 0, column = 0, padx = 5, pady = 5)
    
def getProjectInfo():
    
    global transaction
    transaction = 'Quote'
    
    global date
    date = str(DateE.get())
    
    global QuoteRef
    QuoteRef = ReferenceE.get()
    
    global QuoteMessage
    QuoteMessage  = MessageE.get()
    
    global CustomerCode
    CCode = CCodeE.get()
    Customerarray = CCode.split('-')
    CustomerCode = Customerarray[0]
    
    global termname
    termname = 'CASH'
    
    global state
    state = 'Pending'
    
    global WarehouseID
    WarehouseID = 'Lumax - Olifantsfontein'
    
    global Unit
    Unit = 'Each'
    
    global DepartmentID
    DepartmentID = 'GroundMounting'
    
    global Sodcust
    Sodcust = 'CASH'
    
    newWindow.destroy()
    
def CreateSageImport():
    global df
    
    # Step 1: Read the quote template CSV file into a DataFrame
    template_file = 'import template.csv'
    template_df = pd.read_csv(template_file)

    # Display the template DataFrame
    #print("Template DataFrame:")
    #print(template_df.head())

    # Step 2: Assume you have a quote DataFrame with new data
    # Example quote DataFrame (replace this with your actual quote DataFrame)
    #quote_file = 'test.xlsx'
    #quote_df = pd.read_excel(quote_file)
    quote_df = df
    quote_df = quote_df.loc[:, ['Code', 'Description', 'Quantity', 'Discount Price']]
    quote_df.rename(columns={'Code': 'ITEMID', 'Description': 'ITEMDESC', 'Quantity': 'QUANTITY', 'Discount Price': 'PRICE'}, inplace=True)
                                            
    # Display the quote DataFrame
    #print("\nQuote DataFrame:")
    #print(quote_df.head())

    # Step 3: Create a list of dictionaries to represent rows for the merged DataFrame
    merged_data = []

    # Add the first row of the template DataFrame as the header row in the merged DataFrame
    merged_data.append(dict(zip(template_df.columns, template_df.iloc[0])))

    # Iterate over each row in the quote DataFrame and map it to the template columns
    for _, quote_row in quote_df.iterrows():
        # Create a dictionary to hold data for the new row
        new_row = {}

        # Map the quote data to the corresponding template columns
        for col in quote_df.columns:
            if col in template_df.columns:
                new_row[col] = quote_row[col]  # Assign quote data to the corresponding template column

        # Append the new row dictionary to the list
        merged_data.append(new_row)

    # Create the merged DataFrame directly from the list of dictionaries
    global merged_df
    merged_df = pd.DataFrame(merged_data, columns=template_df.columns)

    #Updating the line 1 items:
    merged_df.at[1, 'TRANSACTIONTYPE'] = transaction

    merged_df.at[1, 'DATE'] = date

    merged_df.at[1, 'GLPOSTINGDATE'] = date

    merged_df.at[1, 'CUSTOMER_ID'] = CustomerCode

    merged_df.at[1, 'TERMNAME'] = termname

    merged_df.at[1, 'REFERENCENO'] = QuoteRef

    merged_df.at[1, 'MESSAGE'] = QuoteMessage

    merged_df.at[1, 'STATE'] = state

    for i in range(1, (len(merged_df.index) - 1)):
        merged_df.at[i, 'LINE'] = i
        merged_df.at[i, 'WAREHOUSEID'] = WarehouseID
        merged_df.at[i, 'UNIT'] = "Each"
        merged_df.at[i, 'DEPARTMENTID'] = DepartmentID
        merged_df.at[i, 'LOCATIONID'] = "100 - Lumax"
        merged_df.at[i, 'SODOCUMENTENTRY_CUSTOMERID'] = Sodcust
        
    # Display the merged DataFrame
    #print("\nMerged DataFrame:")
    #print(merged_df)

    # Step 4: Save the merged DataFrame to a new CSV file
    Save_CSV()

def Save_CSV():
    global merged_df
    file = filedialog.asksaveasfilename(defaultextension = ".csv")
    merged_df.to_csv(str(file), index=False)
    label_file.config(text = "File saved")
    
root = tk.Tk()
root.geometry("1380x750")
root.title("Economax Quote Tool (Use at your own risk)")

InputFrame = tk.LabelFrame(root, text = "Table Data Entry: ")
InputFrame.pack(side = "top", fill = "x")
#InputFrame.place(height = 400, width = 1380)

DispFrame = tk.LabelFrame(root, text = "Calculated Quote: ")
DispFrame.pack(expand = True,fill = "both")
#DispFrame.place(height = 350, width = 1380, rely = 0.525, relx = 0)

LoadPricesB = tk.Button(InputFrame, text = "Load current Prices", command = lambda: Load_excel_data())
LoadPricesB.grid(row = 1, column = 1, padx = 5, pady = 5)

CustomerInfoB = tk.Button(InputFrame, text = "Enter Project Info", command = lambda: ProjectInfo())
CustomerInfoB.grid(row = 1, column = 3, padx = 5, pady = 5)

label_file = ttk.Label(InputFrame, text = "No file selected")
label_file.grid(row = 1, column = 2, padx = 5, pady = 5)

debugLabel = tk.Label(InputFrame, text = "Lol")
debugLabel.grid(row = 1, column = 5, padx = 5, pady = 5)

SupportLabel = tk.Label(InputFrame, text = "Please email kavish@lumaxenergy.com to report any bugs.")
SupportLabel.grid(row = 1, column = 6, padx = 5, pady = 5)

TableNumberLabel = tk.Label(master = InputFrame, text = "Number of Tables:")
TableNumberLabel.grid(row = 2, column = 1, padx = 5, pady = 5)
TableNumberE = tk.Entry(InputFrame)
TableNumberE.grid(row = 2, column = 2, padx = 5, pady = 5)

MountLabel = tk.Label(InputFrame, text = "Rafter Type Choice:")
MountLabel.grid(row = 2, column = 3, padx = 5, pady = 5)
MountVar = tk.StringVar()
MountStr = ['R138 Rafter', 'R162 Rafter']
MountVar.set(MountStr[0])
MountOp = tk.OptionMenu(InputFrame, MountVar, *MountStr, command = rafter_choice_selected)
MountOp.grid(row = 2, column = 4, padx = 5, pady = 5)


#VertPan = ['1','2','3', '4', '6']
#VertVar = tk.StringVar()
#VertVar.set(VertPan[0])
VertPanelLabel = tk.Label(InputFrame, text = "Table Width (no. of panels):")
VertPanelLabel.grid(row = 4, column = 1, padx = 5, pady = 5)
VertPanelE = tk.Entry(InputFrame)
VertPanelE.grid(row = 4, column = 2, padx = 5, pady = 5)

OrientationLabel = tk.Label(InputFrame, text = "Panel Orientation:")
OrientationLabel.grid(row = 3, column = 1, padx = 5, pady = 5)
OrientationVar = tk.StringVar()
OrientationList = ['Portrait', 'Landscape']
OrientationVar.set(OrientationList[1])
OrientationOp = tk.OptionMenu(InputFrame, OrientationVar, *OrientationList)
OrientationOp.grid(row = 3, column = 2, padx=5, pady=5)

HorPanelLabel = tk.Label(InputFrame, text = "Table Length (no. of panels):")
HorPanelLabel.grid(row = 5, column = 1, padx = 5, pady = 5)
HorPanelE = tk.Entry(InputFrame)
HorPanelE.grid(row = 5, column = 2, padx = 5, pady = 5)

DiscountLabel = tk.Label(InputFrame, text = "Customer Discount [%]:")
DiscountLabel.grid(row = 7, column = 1, padx = 5, pady = 5)
DiscountE = tk.Entry(InputFrame)
DiscountE.grid(row = 7, column = 2, padx = 5, pady = 5)

PanelWidthLabel = tk.Label(InputFrame, text = "Width of the selected panels:")
PanelWidthLabel.grid(row = 3, column = 3, padx = 5, pady = 5)
PanelWidthE = tk.Entry(InputFrame)
PanelWidthE.grid(row = 3, column = 4, padx = 5, pady = 5)

PanelLengthLabel = tk.Label(InputFrame, text = "Length of the selected panels:")
PanelLengthLabel.grid(row = 4, column = 3, padx = 5, pady = 5)
PanelLengthE = tk.Entry(InputFrame)
PanelLengthE.grid(row = 4, column = 4, padx = 5, pady = 5)

AngleLabel = tk.Label(InputFrame, text = "Angle (degrees):")
AngleLabel.grid(row = 5, column = 3, padx = 5, pady = 5)
AngleE = tk.Entry(InputFrame)
AngleE.grid(row = 5, column = 4, padx = 5, pady = 5)

GroundClearanceLabel = tk.Label(InputFrame, text = "Ground Clearance:")
GroundClearanceLabel.grid(row = 6, column = 3, padx = 5, pady = 5)
GroundClearanceE = tk.Entry(InputFrame)
GroundClearanceE.grid(row = 6, column = 4, padx = 5, pady = 5)

SDLabel = tk.Label(InputFrame, text = "Single or Double-access:")
SDLabel.grid(row = 7, column = 3, padx = 5, pady = 5)
SDVar = tk.StringVar()
SDList = ['Single-access', 'Double-access']
SDVar.set(SDList[0])
SDOp = tk.OptionMenu(InputFrame, SDVar, *SDList)
SDOp.grid(row = 7, column = 4, padx = 5, pady = 5)

KBLabel = tk.Label(InputFrame, text = "Number of knee-braces per support:")
KBLabel.grid(row = 3, column = 5, padx = 5 , pady = 5)
KBVar = tk.StringVar()
KBList = ['0', '1', '2', '3', '4', '5', '6']
KBVar.set(KBList[0])
KBOp = tk.OptionMenu(InputFrame, KBVar, *KBList)
KBOp.grid(row = 3, column = 6, padx = 5, pady = 5)


MemberList()
RaftSLabel = tk.Label(InputFrame, text = "Rafter Splice?")
RaftSLabel.grid(row = 4, column = 5, padx = 5, pady = 5)
RaftSVar = tk.StringVar()
RaftSList = ['No', 'Yes']
RaftSVar.set(RaftSList[0])
RaftSOp = tk.OptionMenu(InputFrame, RaftSVar, *RaftSList)
RaftSOp.grid(row = 4, column = 6, padx=5, pady=5)

FRaftOvLabel = tk.Label(InputFrame, text = "Front Rafter Overhang [mm]:") 
FRaftOvLabel.grid(row = 5, column = 5, padx = 5, pady = 5)
FRaftOvE = tk.Entry(InputFrame)
FRaftOvE.grid(row = 5, column = 6, padx = 5, pady = 5)

RRaftOvLabel = tk.Label(InputFrame, text = "Rear Rafter Overhang [mm]:")
RRaftOvLabel.grid(row = 6, column = 5, padx = 5, pady =5)
RRaftOvE = tk.Entry(InputFrame)
RRaftOvE.grid(row = 6, column = 6, padx = 5, pady = 5)

SteelRateLabel = tk.Label(InputFrame, text = 'Steel Rate [R/kg]:')
SteelRateLabel.grid(row = 7, column = 5, padx = 5, pady = 5)
SteelRateE = tk.Entry(InputFrame)
SteelRateE.grid(row = 7, column = 6, padx = 5, pady = 5)

ROHLabel = tk.Label(InputFrame, text = "Please select a total panel overhang on rafter:")
ROHLabel.grid(row = 6, column = 1, padx = 5, pady = 5)
var = tk.StringVar()
RaftOvList = ['600mm', '800mm']
var.set(RaftOvList[0])
RaftOvOp = tk.OptionMenu(InputFrame, var, *RaftOvList)
RaftOvOp.grid(row = 6, column = 2, padx = 5, pady = 5)

SSmallsLabel = tk.Label(InputFrame, text = "Extra Fasteners and Clamps Percentage:")
SSmallsLabel.grid(row = 8, column = 1, padx = 5, pady = 5)
global SSmallsVar
SSmallsVar = tk.StringVar()
SSmallsList = ['2%', '5%', '10%']
SSmallsVar.set(SSmallsList[0])
SSmallsOp = tk.OptionMenu(InputFrame, SSmallsVar, *SSmallsList)
SSmallsOp.grid(row = 8, column = 2, padx = 5, pady = 5)

ConSmallsLabel = tk.Label(InputFrame, text = "Extra Connectors(TTC's, FP's) Percentage:")
ConSmallsLabel.grid(row = 8, column = 3, padx = 5, pady = 5)
global ConSmallsVar
ConSmallsVar = tk.StringVar()
ConSmallsList = ['0%', '5%', '10%']
ConSmallsVar.set(ConSmallsList[0])
ConSmallsOp = tk.OptionMenu(InputFrame, ConSmallsVar, *ConSmallsList)
ConSmallsOp.grid(row = 8, column = 4, padx = 5, pady = 5)

SuppSmallsLabel = tk.Label(InputFrame, text = "Extra Supports Percentage:")
SuppSmallsLabel.grid(row = 8, column = 5, padx = 5, pady = 5)
global SuppSmallsVar
SuppSmallsVar = tk.StringVar()
SuppSmallsList = ['0%', '5%', '10%']
SuppSmallsVar.set(SuppSmallsList[0])
SuppSmallsOp = tk.OptionMenu(InputFrame, SuppSmallsVar, *SuppSmallsList)
SuppSmallsOp.grid(row = 8, column = 6, padx = 5, pady = 5)

CalcRaftB = tk.Button(InputFrame, text = "Calculate Rafter Length", command = lambda: Calculations())
CalcRaftB.grid(row = 9, column = 1, padx = 5, pady = 5)
CalcRaftLabel = tk.Label(InputFrame, text = " ")
CalcRaftLabel.grid(row = 9, column = 2, padx = 5, pady = 5)

global rafter_length_options
rafter_length_options = {
    "R138 Rafter": R138RafterList,
    "R162 Rafter": R162RafterList
}
RafterChoiceLabel = tk.Label(InputFrame, text = "Please select a Rafter Length:")
RafterChoiceLabel.grid(row = 10, column = 1, padx = 5, pady = 5)
global RaftVar
RaftVar = tk.StringVar()
#RaftStr = ['3400', '3600', '3800', '4000', '4200', '4400', '5400', '5600', '6200']
RaftStr = Rafterdf['Rafter Code'].tolist()
RaftVar.set(RaftStr[0])
RafterChoiceOp = tk.OptionMenu(InputFrame, RaftVar, *RaftStr)
RafterChoiceOp.grid(row = 10, column = 2, padx = 5, pady = 5)

CalcPurlLabel = tk.Label(InputFrame, text = "Calculated Purlin Length")
CalcPurlLabel.grid(row = 11, column = 1, padx = 5, pady = 5)

PurlinLabel = tk.Label(InputFrame, text = "Supplied Purlin Length:")
PurlinLabel.grid(row = 11, column = 2, padx = 5, pady = 5)

SupportSLabel = tk.Label(InputFrame, text = "Support Spacing")
SupportSLabel.grid(row = 11, column = 3, padx = 5, pady = 5)

SupportLegsLabel = tk.Label(InputFrame, text = "Support Legs")
SupportLegsLabel.grid(row = 11, column = 4, padx = 5, pady = 5)

OHangLabel = tk.Label(InputFrame, text = "Overhang")
OHangLabel.grid(row = 11, column = 5, padx = 5, pady = 5)

TotalPriceLabel = tk.Label(InputFrame, text = "Total Price of the quote:")
TotalPriceLabel.grid(row = 11, column = 6, padx = 5, pady = 5)

CalcQuoteB = tk.Button(InputFrame, text = "Calculate Quote", command = lambda: FinishCalc())
CalcQuoteB.grid(row = 12, column = 1, padx = 5, pady = 5)

DispQuoteB = tk.Button(InputFrame, text = "Display Quote", command = lambda: Refresh())
DispQuoteB.grid(row = 12, column = 2, padx = 5, pady = 5)

ExportB = tk.Button(InputFrame, text = "Export Quote", command = lambda: Save_Excel())
ExportB.grid(row = 12, column = 3, padx = 5, pady = 5)

SageB = tk.Button(InputFrame, text = "Create Sage Import", command = lambda: CreateSageImport())
SageB.grid(row = 12, column = 4, padx = 5, pady = 5)

# Treeview Widget
tv1 = ttk.Treeview(DispFrame)
tv1.place(relheight=1, relwidth=1)

treescrolly = tk.Scrollbar(DispFrame, orient = "vertical", command=tv1.yview)
treescrollx = tk.Scrollbar(DispFrame, orient = "horizontal", command = tv1.xview)
tv1.configure(xscrollcommand = treescrollx.set, yscrollcommand = treescrolly.set)
treescrollx.pack(side = "bottom", fill = "x")
treescrolly.pack(side = "right", fill = "y")

# Add weights to the grid rows and columns
# Changing the weights will change the size of the rows/columns relative to each other
DispFrame.grid_rowconfigure(0, weight=1)
DispFrame.grid_rowconfigure(1, weight=1)
DispFrame.grid_columnconfigure(0, weight=1)
DispFrame.grid_columnconfigure(1, weight=1)

root.mainloop()