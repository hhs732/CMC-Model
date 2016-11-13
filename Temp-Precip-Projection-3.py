import numpy as np
import xlrd

#%%
# *********** Import & Read Data ********** #
Data1 = xlrd.open_workbook('Data.xlsx')
nsheet = Data1.nsheets                
sheetnames = Data1.sheet_names()

def f(x,y,z):
        sheetz = Data1.sheet_by_index(z)
        return np.array(sheetz.cell(x,y).value)
 #%%       
dic = {}
for k in range (0,nsheet):
    sheetnamez=sheetnames[k]
    sheetz = Data1.sheet_by_index(k)
    #fz = f(k)
    R = np.array(sheetz.nrows)
    C = np.array(sheetz.ncols)
    
    M = np.empty((R,C))
    dt = np.empty((R,C))
    
    for i in range(R): 
        for j in range(C): 
            #print (i,j)
            h = f(i,j,k)
            #print (z)
            dt = h.dtype
            #print (dt)
            if dt!='float64':
                M[i,j]=np.array(-999999)
            else: M[i,j]=h
            np.set_printoptions(precision=3, suppress=True)
            dic[sheetnamez]=M
#print (M)
#print (dic) 
#%%
# ************** Input Data ************** #
Input = dic.get('InputData')
# ************** 5 Modules *************** #
Mod1 = dic.get('NPHX')      # ****NPHX**** #  
Mod2 = dic.get('Jet120W')   # ***Jet120W** #
Mod3 = dic.get('EUHX')      # ****EUHX**** #
Mod4 = dic.get('NHT')       # ****MNHT**** #
Mod5 = dic.get('ITC90W')    # ***ITC90W*** #
#print(Input.item(1,1))

# ************* Constants **************** #
CT = -28.4578180051601
XCFNHT = 2.97759352945703
XCFJet120W = -0.246170299640341

CP = 2.2164393012202
XCEUHX = -0.0271538631833865	
XCITC = 0.22227876348189	  #XCITC90W
XCFcnJet = -0.669882303575894
XCFcnEUHX = 0.308785154938417
XCFcnITC = 0.6743742667085
XCFcnNPHX =	-0.392794315800777
LatJet = 28.9 
WidJet = 0.3
LatEUHX = 45.1
WidEUHX = 0.5
LatITC = 14
WidITC = 0.3
LatNPHX = 38
WidNPHX = 0.5

# *********** Input Variables *********** #
Tobs = Input[5:17 ,4]
NHTmonth = Mod4[4, 2:] 
NHTyear = Mod4[5:, 2:]
Jet120Wmonth = Mod2[3, 2:]
Jet120Wyear = Mod2[4:, 2:]

Pobs = Input[5:17 ,1]
NPHXmonth = Mod1[2, 2:]
NPHXyear = Mod1[3:, 2:]
EUHXmonth = Mod3[3, 2:]
EUHXyear = Mod3[4:, 2:]
ITCmonth = Mod5[12, 2:]
ITCyear = Mod5[13:, 2:]
#%%
# *********** Temperature Projection *********** #
#def T(NHTmonthy,XCFNHT1,XCFJet120W1,Jet120Wmonthy,CT1,Tobsy,NHTyearxy,Jet120Wyearxy):
#    TRawCalcNowy = (XCFNHT1 * NHTmonthy)+(XCFJet120W1 * Jet120Wmonthy) + CT1
#    out1=CT + Tobsy - TRawCalcNowy + (XCFNHT1*NHTyearxy) + (XCFJet120W1*Jet120Wyearxy)
#    return out1 

HTemp = np.empty((len(NHTyear),len(NHTyear.T)))
for p in range (len(NHTyear.T)):     #12
    print p
    for q in range (len(NHTyear)):   #400
        #MeanTY = T(NHTmonth[p],XCFNHT,XCFJet120W,Jet120Wmonth[p],CT,Tobs[p],NHTyear[q,p],Jet120Wyear[q,p])
        TRawCalcNowy = (XCFNHT * NHTmonth[p])+(XCFJet120W * Jet120Wmonth[p]) + CT
        HTemp[q,p]=CT + Tobs[p] - TRawCalcNowy + (XCFNHT*NHTyear[q,p]) + (XCFJet120W*Jet120Wyear[q,p])
        np.set_printoptions(precision=2, suppress=True)
print(HTemp)
# ********** Precipitation Projection ********** #
PPTCRm = np.power(Pobs,0.33333)

def PRaw(Y):
    FcnJetmy = np.exp(-0.5*np.power(((LatJet-Jet120Wmonth[Y])/WidJet),2))
    FcnEUHXmy= np.exp(-0.5*np.power(((LatEUHX-EUHXmonth[Y])/WidEUHX),2)) #FcnHigh
    FcnITCmy = np.exp(-0.5*np.power(((LatITC-ITCmonth[Y])/WidITC),2))
    FcnNPHXmy = np.exp(-0.5*np.power(((LatNPHX-NPHXmonth[Y])/WidNPHX),2)) #Other
    out1=np.power((CP + XCEUHX * EUHXmonth[Y] + XCITC * ITCmonth[Y] + XCFcnJet * FcnJetmy + XCFcnEUHX * FcnEUHXmy + XCFcnITC * FcnITCmy + XCFcnNPHX * FcnNPHXmy),3)    
    return out1 

def P(X,Y):
    FcnJetyxy = np.exp(-0.5*np.power(((LatJet-Jet120Wyear[X,Y])/WidJet),2))
    FcnEUHXyxy = np.exp(-0.5*np.power(((LatEUHX-EUHXyear[X,Y])/WidEUHX),2)) #FcnHigh
    FcnITCyxy = np.exp(-0.5*np.power(((LatITC-ITCyear[X,Y])/WidITC),2))
    FcnNPHXyxy = np.exp(-0.5*np.power(((LatNPHX-NPHXyear[X,Y])/WidNPHX),2)) #Other
    out1=(np.power((CP + (XCEUHX * EUHXyear[X,Y]) + (XCITC * ITCyear[X,Y]) + (XCFcnJet * FcnJetyxy) + (XCFcnEUHX * FcnEUHXyxy) + (XCFcnITC * FcnITCyxy) + (XCFcnNPHX * FcnNPHXyxy)),3)) * (Pobs[Y] / PRaw(Y))    
    return out1 
    
HPrecip = np.empty((len(NHTyear),len(NHTyear.T)))
#%%
for p in range (len(NHTyear.T)):     #12
    for q in range (len(NHTyear)):   #400
        PrecipY = P(q,p) 
        HPrecip[q,p]=PrecipY
        np.set_printoptions(precision=2, suppress=True)
print(HPrecip)



























