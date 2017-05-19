###############################################################################
#           Automated Test of CCSI Moving Bed Simulator
#           Runs on Aspen Customer Modeler 7.3
#           Uses Python 2.7, pywinauto 0.4.3
#           Gregory Pope, LLNL
#           February 5, 2013
#           For Aceptance Testing Moving Bed Reactor Model v1.2.1a
#           V1.0
###############################################################################

import pywinauto
import time

def ChangeMouse(script):
     #Assumes Ghost mouse is visable on desktop and scripts are in My Documents folder
     #Change(or load intital) mouse scripts into Ghost Mouse
     #Allows Ghost Mouse to change scripts during automated test run to create numerous mouse behaviors 
     print('Change Mouse Scripts')
     w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'GhostMouse 3.2', class_name='AutoIt v3 GUI')[0])
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     
     window.MenuItem(u'&File').Click()
     window.MenuItem(u'&File->&Open').Click()
     w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Open file...', class_name='#32770')[0])
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     ctrl = window['ComboBox']
     ctrl.SetFocus()
     ctrl.ClickInput()
     ctrl.TypeKeys(script)
     ctrl.TypeKeys('{ENTER}')
     return

def SetMouseSpeed(speed):
     #Assumes Ghost mouse is visable on desktop
     #Change(or load intital) mouse speed setting into Ghost Mouse (speed is non-zero integer -10 to 10, 10 = 10 times faster, -10 1/10th speed)
     #Allows Ghost Mouse to change speeds during automated test run to create numerous mouse behaviors 
     w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'GhostMouse 3.2', class_name='AutoIt v3 GUI')[0])
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     window.MenuItem(u'&Options').Click()
     window.MenuItem(u'&Options->&Playback').Click()
     window.MenuItem(u'&Options->&Playback->Spee&d').Click()
     w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Speed setting', class_name='AutoIt v3 GUI')[0])
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     ctrl = window['Trackbar']
     ctrl.SetFocus()
     if (speed == 10): ctrl.TypeKeys('{END}')
     elif (speed == -10): ctrl.TypeKeys('{HOME}')
     elif (speed == 1): ctrl.TypeKeys('{END}'+'{LEFT 9}')
     elif (speed == 2): ctrl.TypeKeys('{END}'+'{LEFT 8}')
     elif (speed == 3): ctrl.TypeKeys('{END}'+'{LEFT 7}')
     elif (speed == 4): ctrl.TypeKeys('{END}'+'{LEFT 6}')
     elif (speed == 5): ctrl.TypeKeys('{END}'+'{LEFT 5}')
     elif (speed == 6): ctrl.TypeKeys('{END}'+'{LEFT 4}')
     elif (speed == 7): ctrl.TypeKeys('{END}'+'{LEFT 3}')
     elif (speed == 8): ctrl.TypeKeys('{END}'+'{LEFT 2}')
     elif (speed == 9): ctrl.TypeKeys('{END}'+'{LEFT 1}')
     elif (speed == -9): ctrl.TypeKeys('{HOME}'+'{RIGHT 1}')
     elif (speed == -8): ctrl.TypeKeys('{HOME}'+'{RIGHT 2}')
     elif (speed == -7): ctrl.TypeKeys('{HOME}'+'{RIGHT 3}')
     elif (speed == -6): ctrl.TypeKeys('{HOME}'+'{RIGHT 4}')
     elif (speed == -5): ctrl.TypeKeys('{HOME}'+'{RIGHT 5}')
     elif (speed == -4): ctrl.TypeKeys('{HOME}'+'{RIGHT 6}')
     elif (speed == -3): ctrl.TypeKeys('{HOME}'+'{RIGHT 7}')
     elif (speed == -2): ctrl.TypeKeys('{HOME}'+'{RIGHT 8}')
     else: print ('speed must be non-zero integer -10 to 10')
     ctrl = window['Ok']
     ctrl.Click()
     return

#File needed to do simualtion in Aspen Custom Modeler/AM_Untitled
Needed_file='MovingB_v1.2.1a.acmf'
#assure we do not type too fast over remore set up
keywait =.33

pwa_app = pywinauto.application.Application()

# Assure Ghost Mouse help application in tray using file drag_moving_bed.rms
raw_input('Assure Ghost Mouse app. on Desk Top, hit enter')


# Start Application
pwa_app.start_("C:\Program Files (x86)\AspenTech\AMSystem V7.3\Bin\AspenModeler.exe")
print ('Started Application')


#Dismiss Registration if present
print ('Dismiss Registration')
if (pywinauto.timings.WaitUntilPasses(5,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Aspen Technology Product Registration', class_name='#32770')[0])):
     w_handle = pywinauto.findwindows.find_windows(title=u'Aspen Technology Product Registration', class_name='#32770')[0]
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     ctrl = window['Button2']
     ctrl.Click()


#Select and Configure Physical Properties
print ('Select Physical Properties')
w_handle = pywinauto.timings.WaitUntilPasses(10,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.Maximize()
window.TypeKeys('%t')
window.TypeKeys('g')

w_handle = pywinauto.timings.WaitUntilPasses(8,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Physical Properties Configuration', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['RadioButton']
ctrl.Click()
ctrl = window['Button2']
ctrl.Click()

print ('Select Chemicals')

w_handle = pywinauto.timings.WaitUntilPasses(10,0.5,lambda:pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Components Specifications - Data Browser]')[0])
window = pwa_app.window_(handle=w_handle)
window.edit.TypeKeys('CO2\r') # for carbon dioxide
window.edit.TypeKeys('H2O\r') # for water
window.edit.TypeKeys('N2\r')  # for Nitrogen

w_handle = pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Components Specifications - Data Browser]')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#Tools->Next
window.TypeKeys('%t')
window.TypeKeys('n')

# Select Property Method
print ('Select Property Method')
ctrl = window['AfxOleControl9023']
ctrl.SetFocus()
ctrl = window['ComboBox14']
ctrl.Click()
ctrl.Select('PENG-ROB')
ctrl.Click()

#Next Step
#Tools->Next
window.TypeKeys('%t')
window.TypeKeys('n')

window.TypeKeys('%t')
window.TypeKeys('n')

w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Required Properties Input Complete', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Required PROPS Input Complete', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Economic Analysis', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Close']
ctrl.Click()

w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda:pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Control Panel]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#Tools->Next
window.TypeKeys('%t')
window.TypeKeys('n')


w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Properties Plus Complete', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


print ('Exit Properties')

w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda:pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Control Panel]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#File->Save
window.TypeKeys('%f')
window.TypeKeys('s')
time.sleep(.5)
# File->Exit
window.edit.TypeKeys("%f")
window.edit.TypeKeys("x")


#This code needed first time and window not disabled TODO 

#w_handle = pywinauto.findwindows.find_windows(title=u'Aspen Properties', class_name='#32770')[0]
#window = pwa_app.window_(handle=w_handle)
#window.SetFocus()
# Don't display anymore
#ctrl = window['&No']
#ctrl.Click()


w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Physical Properties Configuration', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

time.sleep(1)

print('Now move over chemicals')
#Select the Components List view in All Items
w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['AfxMDIFrame902']
ctrl.SetFocus()
ctrl = window['TreeView']
ctrl.ClickInput()
ctrl.ClickInput()

window.TypeKeys('{PGUP}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{TAB}')
#Works in both three across or two across width
window.TypeKeys('{RIGHT}')
window.TypeKeys('{RIGHT}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{ENTER}')

           
#Move the chemicals to the active window
w_handle = pywinauto.timings.WaitUntilPasses(4,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Build Component List - Default', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['List1']
ctrl.SetFocus()
ctrl.Select(0)
ctrl = window['>']
ctrl.Click()
ctrl = window['List1']
ctrl.SetFocus()
ctrl.Select(0)
ctrl = window['>']
ctrl.Click()
ctrl = window['List1']
ctrl.SetFocus()
ctrl.Select(0)
ctrl = window['>']
ctrl.Click()
ctrl = window['OK']
ctrl.Click()

print('Get the needed file')
#Import in the needed file
w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#File -> Import
window.TypeKeys('%f')
window.TypeKeys('i')


w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Open', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['ComboBox']
ctrl.SetFocus()
ctrl = window['Edit']
ctrl.SetFocus()

#file name below may change, see note at top

window.edit.TypeKeys(Needed_file)
ctrl = window['&Open']
ctrl.Click()

print('Get the Adsorber model')
#Get the Adsorber model
w_handle = pywinauto.timings.WaitUntilPasses(6,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Exploring - Simulation']
ctrl.SetFocus()
time.sleep(.5)
#Narrow width for so upcoming ghost mouse script will work
ctrl.MoveWindow(x=0, y=0, width = 200)
ctrl = window['AfxMDIFrame902']
ctrl.SetFocus()
#Assure depth of window so mouse will work
ctrl.MoveWindow(x=0, y=0, height = 238)
ctrl = window['TreeView']
ctrl.SetFocus()
window.TypeKeys('{DOWN}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{TAB}')
window.TypeKeys('{RIGHT}')
window.TypeKeys('{ENTER}')
window.TypeKeys('{TAB}')
window.TypeKeys('{TAB}')
window.TypeKeys('{DOWN}')
#Assure depth of window so mouse will work
ctrl = window['AfxMDIFrame903']
ctrl.SetFocus()
ctrl.MoveWindow(x=0, y=238, height = 610)

w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.ClickInput()

#Load mouse sript
ChangeMouse('drag_moving_bed.rms')
SetMouseSpeed(1)

#Run Ghost Mouse Script to drag Model to Process Flowsheet (Control F2)
ctrl.TypeKeys('^{F2}')



#Dismiss the done Dialog Box
print ('Dismiss the done dialog box')
w_handle = pywinauto.timings.WaitUntilPasses(50,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Scripting', class_name='#32770')[0])                                          
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

#Run the initial simulation
print ('Run the initial simulation')

w_handle = pywinauto.timings.WaitUntilPasses(10,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('{F5}')



#Wait for and then dismiss the done dialog box up to 20 seconds
w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda: pywinauto.findwindows.find_windows(title= u'Run complete', class_name='#32770')[0])
print ('Dismiss the done dialog box')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

#Start at Top of Table List (Hx)
print ('Start at Top of Table List (Hx)')
w_handle = pywinauto.timings.WaitUntilPasses(10,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['B2.Config Table']
ctrl.SetFocus()
ctrl.ClickInput()
window.TypeKeys('{PGUP}'+'{DOWN}')

#Config Values:
print ('Fill Adsorber Configure Table')
Hx= 6.0
Dx= 9.0
e=  0.7
Dtube=.02
Tube_N= 4000

window.TypeKeys(str(Hx)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Dx)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(e)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Dtube)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Tube_N)+'{RIGHT}')
time.sleep(keywait)
window.TypeKeys('%{DOWN}'+'{PGUP}')
time.sleep(keywait)
window.TypeKeys('{LEFT}')
time.sleep(keywait)
window.TypeKeys ('{DOWN 12}')
time.sleep(keywait)
#select upward
window.TypeKeys('%{DOWN}'+'{PGDN}')


#Run the first simulation
print ('Run the first simulation')
w_handle = pywinauto.timings.WaitUntilPasses(15,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
time.sleep(.5)
window.TypeKeys('{F5}')


#Wait for and then dismiss the done dialog box up to 20 seconds
w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda: pywinauto.findwindows.find_windows(title= u'Run complete', class_name='#32770')[0])
print ('Dismiss the done dialog box')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


#Start at Top of Table List (GasIN.F)
print ('Start at Top of Table List (GasIN.F)')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['B1.Inlets Table']
ctrl.SetFocus()
ctrl.ClickInput()
window.TypeKeys('{PGUP}')

#Config Values for Adsorber:
print ('Fill Inlets Configure Table')
GasIN_F= 7000
GasIN_P='Free'
GasIN_T= 45
GasIN_zCO2=  0.13
GasIN_zH2O=  0.07
GasIN_zN2=  0.8
SolidIN_Fm= 470000
SolidIN_T= 100
SolidIN_wBic= .05
SolidIN_wCar= .45
SolidIN_wH2O= .53
TubeIN_F=22000
TubeIN_P=1.2
TubeIN_T=33
TubeIN_zCO2=0
TubeIN_zH2O=1
TubeIN_zN2=0

GasOUT_P = 1

window.TypeKeys(str(GasIN_F)+'{RIGHT}'+'{DOWN}')
time.sleep(.25)
window.TypeKeys('%{DOWN}'+'{PGUP}')
time.sleep(keywait)
window.TypeKeys('{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_T)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_zCO2)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_zH2O)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_zN2)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIN_Fm)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIN_T)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIN_wBic)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIN_wCar)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIN_wH2O)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(TubeIN_F)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(TubeIN_P)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(TubeIN_T)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(TubeIN_zCO2)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(TubeIN_zH2O)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(TubeIN_zN2)+'{DOWN}')
time.sleep(keywait)


print ('Change Outlets Configure Table')
ctrl = window['B1.Outlets Table']
ctrl.SetFocus()

#Move Outlets Window onto Screen
ctrl.MoveWindow (840,05)
ctrl.ClickInput()
window.TypeKeys('{PGUP}'+'{DOWN}'+str(GasOUT_P)+'{RIGHT}')
time.sleep(keywait)
window.TypeKeys('%{DOWN}'+'Fixed'+'{DOWN}')
time.sleep(keywait)


#Run the second simulation
print ('Run the second simulation')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('{F5}')


#Wait for and then dismiss the done dialog box up to 25 seconds
w_handle = pywinauto.timings.WaitUntilPasses(25,0.5,lambda: pywinauto.findwindows.find_windows(title= u'Run complete', class_name='#32770')[0])
print ('Dismiss the done dialog box')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


print ('Run the Initialization Script')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.Maximize()
#Force the Icon to the middle of the screen
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE - [Process Flowsheet Window]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('d')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Find Object', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#Find the icon
ctrl = window['ListBox']
ctrl.Click()
ctrl = window['Find']
ctrl.Click()
window.Close()
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE - [Process Flowsheet Window]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
# Put mouse on center of window also and click to fill in Flowsheet->Srcipts menu
ctrl.ClickInput()
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE - [Process Flowsheet Window]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.Restore()
#Start next run using Flowsheet pulldown
window.TypeKeys('%l')
window.TypeKeys('p')
window.TypeKeys('{ENTER}')

#Wait for and then dismiss the done dialog box up to 120 seconds
w_handle = pywinauto.timings.WaitUntilPasses(120,0.5,lambda: pywinauto.findwindows.find_windows(title= u'Scripting', class_name='#32770')[0])
print ('Dismiss the done dialog box\n\n')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

print ('Test Results Moving Bed Aspen Modeler Simulation '+ Needed_file +' '+ time.asctime())
#Check homotopy correct
#Highlight value to be checked
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['B1.Config Table']
ctrl.SetFocus()
ctrl.ClickInput()
ctrl.TypeKeys('{PGDN}')
time.sleep(keywait)
ctrl.TypeKeys('{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys('{LEFT}')
time.sleep(keywait)
ctrl.TypeKeys('{F2}')
#Make sure cursor hightlights selected value of interest
time.sleep(.5)
ctrl = window['Edit']
time.sleep(keywait)
ctrl.SetFocus()

#Get expected result, compare to actual results

homotopy = ['R_htw','R_hts','R_r1e','R_r2e','R_r3e']

expected= 1.0
i=1
while (i<=5):
     
     #Get value highlighted
     Properties=ctrl.Texts()

     #Convert string to number
     actual= float(Properties[1])

     #Compare
     if (actual == expected):
          print ('Test '+str(i)+' pass '+ homotopy[i-1]+' '+'Actual= '+str(actual)+' '+'Expected= '+str(expected)+ ' Difference= '+str(abs(actual-expected)))
     else:
          print ('Test '+str(i)+' fail '+ homotopy[i-1]+' '+'Actual= '+str(actual)+' '+'Expected= '+str(expected)+ ' Difference= '+str(abs(actual-expected)))
          
     #Highlight next
     ctrl.TypeKeys('{DOWN}'+'{F2}')
     i=i+1
     time.sleep(1)
   
#################################################


print('Get the Regenerator model')
#Get the Regenerator  model
w_handle = pywinauto.timings.WaitUntilPasses(6,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Exploring - Simulation']
ctrl.SetFocus()
time.sleep(.5)
#Narrow width for so upcoming ghost mouse script will work
ctrl.MoveWindow(x=0, y=0, width = 200)
ctrl = window['AfxMDIFrame902']
ctrl.SetFocus()
ctrl = window['TreeView']
ctrl.SetFocus()
window.TypeKeys('{DOWN}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{TAB}')
time.sleep(1)
window.TypeKeys('{RIGHT}')
window.TypeKeys('{ENTER}')
window.TypeKeys('{TAB}')
window.TypeKeys('{TAB}')
window.TypeKeys('{DOWN}')

w_handle = pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.ClickInput()
ctrl.MoveWindow(100,8)


#Run Ghost Mouse Script to drag Model to Process Flowsheet (Control F2)
ctrl.TypeKeys('^{F2}')



#Dismiss the done Dialog Box
print ('Dismiss the done dialog box')
w_handle = pywinauto.timings.WaitUntilPasses(50,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Scripting', class_name='#32770')[0])                                          
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

#Run the initial simulation
print ('Run the initial simulation')

w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('{F5}')



#Wait for and then dismiss the done dialog box up to 20 seconds
w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda: pywinauto.findwindows.find_windows(title= u'Run complete', class_name='#32770')[0])
print ('Dismiss the done dialog box')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

#Start at Top of Table List (Hx)
print ('Start at Top of Table List (Hx)')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['B2.Config Table']
ctrl.SetFocus()
ctrl.ClickInput()
window.TypeKeys('{PGUP}'+'{DOWN}')

#Config Values for regenerator:
print ('Fill Regenerator Configure Table')
Hx= 4.0
Dx= 7.0
e=  0.6
Dtube=.01
Tube_N= 16000

window.TypeKeys(str(Hx)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Dx)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(e)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Dtube)+'{DOWN}')
time.sleep(keywait)
#Change variables for Tube_N and Tube_lp
window.TypeKeys(str(Tube_N)+'{RIGHT}')
time.sleep(keywait)
window.TypeKeys('%{DOWN}'+'fi'+'{ENTER}')
time.sleep(keywait)
time.sleep(.5)
window.TypeKeys('{DOWN}'+'%{DOWN}'+'fr'+'{ENTER}')
time.sleep(keywait)
window.TypeKeys('{LEFT}')
time.sleep(keywait)
window.TypeKeys ('{DOWN 11}')
time.sleep(keywait)
window.TypeKeys('%{DOWN}'+'{PGUP}')
time.sleep(keywait)


#Run the first simulation
print ('Run the first simulation')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
time.sleep(.5)
window.TypeKeys('{F5}')


#Wait for and then dismiss the done dialog box up to 9 seconds
w_handle = pywinauto.timings.WaitUntilPasses(9,0.5,lambda: pywinauto.findwindows.find_windows(title= u'Run complete', class_name='#32770')[0])
print ('Dismiss the done dialog box')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


#Start at Top of Table List (GasIN.F)
print ('Start at Top of Table List (GasIN.F)')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['B2.Inlets Table']
ctrl.SetFocus()
ctrl.ClickInput()
window.TypeKeys('{PGUP}')

#Config Values for Generator:
print ('Fill Inlets Configure Table')
GasIN_F= 1000
#GasIN.P and GasOUT.P varialbes change
GasIN_P='Free'
GasIN_T= 120
GasIN_zCO2=  .00001
GasIN_zH2O=  .99998
GasIN_zN2=   .00001
SolidIN_Fm= 470000
SolidIN_T= 56.7932
SolidIN_wBic= .245181
SolidIN_wCar= 2.00478
SolidIN_wH2O= .800932
TubeIN_F=2850
TubeIN_P=6.895
TubeIN_T=200
TubeIN_zCO2=0
TubeIN_zH2O=1
TubeIN_zN2=0

GasOUT_P = 1


window.TypeKeys(str(GasIN_F)+'{RIGHT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys('%{DOWN}'+'{PGUP}')
time.sleep(keywait)
window.TypeKeys('{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_T)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_zCO2)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_zH2O)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_zN2)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIN_Fm)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIN_T)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIN_wBic)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIN_wCar)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIN_wH2O)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(TubeIN_F)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(TubeIN_P)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(TubeIN_T)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(TubeIN_zCO2)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(TubeIN_zH2O)+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(TubeIN_zN2)+'{DOWN}')
time.sleep(keywait)


print ('Change Outlets Configure Table')
ctrl = window['B2.Outlets Table']
ctrl.SetFocus()

#Move Outlets Window onto Screen
ctrl.MoveWindow (840,05)
ctrl.ClickInput()
window.TypeKeys('{PGUP}'+'{DOWN}'+str(GasOUT_P)+'{RIGHT}')
time.sleep(1)
window.TypeKeys('%{DOWN}'+'Fixed'+'{DOWN}')
time.sleep(3)


#Run the second simulation
print ('Run the second simulation')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('{F5}')


#Wait for and then dismiss the done dialog box up to 25 seconds
w_handle = pywinauto.timings.WaitUntilPasses(25,0.5,lambda: pywinauto.findwindows.find_windows(title= u'Run complete', class_name='#32770')[0])
print ('Dismiss the done dialog box')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


print ('Run the Initialization Script')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.Maximize()
#Force the Icon to the middle of the screen
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE - [Process Flowsheet Window]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('d')

w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Find Object', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#Find the icon
ctrl = window['ListBox']
ctrl.Click()
time.sleep(.5)
#Select B2
ctrl.TypeKeys('{DOWN}')
ctrl = window['Find']
ctrl.Click()
window.Close()
w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE - [Process Flowsheet Window]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
# Put mouse on center of window also and click to fill in Flowsheet->Srcipts menu
ctrl.ClickInput()
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE - [Process Flowsheet Window]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.Restore()
#Start next run using Flowsheet pulldown
window.TypeKeys('%l')
window.TypeKeys('p')
window.TypeKeys('{ENTER}')

#Wait for and then dismiss the done dialog box up to 40 minutes
w_handle = pywinauto.timings.WaitUntilPasses(2400,2,lambda: pywinauto.findwindows.find_windows(title= u'Scripting', class_name='#32770')[0])
print ('Dismiss the done dialog box\n\n')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

print ('Test Results Moving Bed Aspen Modeler Simulation '+ Needed_file +' '+ time.asctime())
#Check homotopy correct
#Highlight value to be checked
w_handle = pywinauto.timings.WaitUntilPasses(2,.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['B2.Config Table']
ctrl.SetFocus()
ctrl.ClickInput()
ctrl.TypeKeys('{PGDN}')
ctrl.TypeKeys('{DOWN}')
ctrl.TypeKeys('{LEFT}')
ctrl.TypeKeys('{F2}')
#Make sure cursor highlights selected value of interest
time.sleep(.5)
ctrl = window['Edit']
time.sleep(keywait)
ctrl.SetFocus()

#Get expected result, compare to actual results

homotopy = ['R_htw','R_hts','R_r1e','R_r2e','R_r3e']

expected= 1.0
i=1
while (i<=5):
     
     #Get value highlighted
     Properties=ctrl.Texts()

     #Convert string to number
     actual= float(Properties[1])

     #Compare
     if (actual == expected):
          print ('Test '+str(i+5)+' pass '+ homotopy[i-1]+' '+'Actual= '+str(actual)+' '+'Expected= '+str(expected)+ ' Difference= '+str(abs(actual-expected)))
     else:
          print ('Test '+str(i+5)+' fail '+ homotopy[i-1]+' '+'Actual= '+str(actual)+' '+'Expected= '+str(expected)+ ' Difference= '+str(abs(actual-expected)))
          
     #Highlight next
     ctrl.TypeKeys('{DOWN}'+'{F2}')
     i=i+1
     time.sleep(1)
   





