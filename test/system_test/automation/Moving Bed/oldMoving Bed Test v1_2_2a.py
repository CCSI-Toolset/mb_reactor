###############################################################################
#           Automated Test of CCSI Moving Bed Simulator v1.2.2a
#           Runs on Aspen Customer Modeler 7.3
#           Uses Python 2.7, pywinauto 0.4.3
#           Gregory Pope, LLNL
#           February 5, 2013
#           For Acceptance Testing Moving Bed Reactor Model v1.2.2a
#           Replaced by v1.2.1a
#           V1.0
###############################################################################

import pywinauto
import time

#File needed to do simualtion in Aspen Custom Modeler/AM_Untitled
Needed_file='MovingB_v1.2.2a.acmf'

pwa_app = pywinauto.application.Application()

# Assure Ghost Mouse help application in tray using file drag_moving_bed.rms
raw_input('Assure Ghost Mouse app. in tray using file drag_moving_bed.rms, hit enter')

# Start Application
pwa_app.start_("C:\Program Files (x86)\AspenTech\AMSystem V7.3\Bin\AspenModeler.exe")
print ('Started Application')


#Dismiss Registration
print ('Dismiss Registration')
w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Aspen Technology Product Registration', class_name='#32770')[0])
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
window.MenuItem(u'&Tools->Confi&gure Properties...\tAlt+F6').Click()

w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Physical Properties Configuration', class_name='#32770')[0])
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
window.MenuItem(u'&Tools->&Next\t F4').Click()

# Select Property Method
print ('Select Property Method')
ctrl = window['AfxOleControl9023']
ctrl.SetFocus()
ctrl = window['ComboBox14']
ctrl.Click()
ctrl.Select('PENG-ROB')
ctrl.Click()

#Next Step
window.MenuItem(u'&Tools->&Next\t F4').Click()
time.sleep (.5)
window.MenuItem(u'&Tools->&Next\t F4').Click()

pwa_app = pywinauto.application.Application()
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
window.MenuItem(u'&Tools->&Next\t F4').Click()

w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Properties Plus Complete', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


print ('Exit Properties')

w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda:pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Control Panel]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.MenuItem(u'&File->&Save\tCtrl+S').Click()
time.sleep(.5)
#window.MenuItem(u'&File->E&xit').Click() does not work I think becasue the Menu is too long and it can not find Exit at the end.
# work around below works
window.edit.TypeKeys("%f")   # alt f = File
window.edit.TypeKeys("x")    #  x = Exit


#This code needed if first time and window not disabled
#w_handle = pywinauto.findwindows.find_windows(title=u'Aspen Properties', class_name='#32770')[0]
#window = pwa_app.window_(handle=w_handle)
#window.SetFocus()
#ctrl = window['&No']
#ctrl.Click()

#time.sleep(1)

w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Physical Properties Configuration', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


print('Now move over chemicals')
#Select the Components List view in All Items
w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Exploring - Simulation']
ctrl.SetFocus()
ctrl = window['AfxMDIFrame902']
ctrl.SetFocus()
ctrl = window['TreeView']
ctrl.SetFocus()
window.TypeKeys('{DOWN}')

#Go to contents of Components List and Select Default
ctrl = window['AfxMDIFrame903']
ctrl.SetFocus()
ctrl = window['ListView']
ctrl.SetFocus()
ctrl.Select(2)
ctrl.ClickInput(coords=(30,100),double='true')
           

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
window.MenuItem(u'&File->&Import Types...').Click()


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


print('Get the model')
#Get the model
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
ctrl.MoveWindow(x=0, y=0, height = 240)
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
ctrl.MoveWindow(x=0, y=240, height = 605)

#Run Ghost Mouse Script (Control F2)
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
w_handle = pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('{F5}')
#window.MenuItem(u'&Run->R&un\tF5').Click()

#Wait and then dismiss dialog box, max wait is 8 seconds before time out. 
w_handle = pywinauto.timings.WaitUntilPasses(8,0.5,lambda: pywinauto.findwindows.find_windows(title= u'Run complete', class_name='#32770')[0])
print ('Dismiss the done dialog box')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

#Start at Top of Table List (Hx)
print ('Start at Top of Table List (Hx)')
w_handle = pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['B1.Config Table']
ctrl.SetFocus()
ctrl.ClickInput()
window.TypeKeys('{PGUP}'+'{DOWN}')

#Config Values:
print ('Fill Configure Table')
Hx= 6.0
Dx= 9.0
e=  0.7
Dtube=.02
Tube_N= 4000

window.TypeKeys(str(Hx)+'{DOWN}')
window.TypeKeys(str(Dx)+'{DOWN}')
window.TypeKeys(str(e)+'{DOWN}')
window.TypeKeys(str(Dtube)+'{DOWN}')
window.TypeKeys(str(Tube_N)+'{RIGHT}')
window.TypeKeys('%{DOWN}'+'{PGUP}')
time.sleep(.5)
window.TypeKeys('{LEFT}')
window.TypeKeys ('{DOWN 11}')
time.sleep(1)
window.TypeKeys('%{DOWN}'+'{PGDN}')
time.sleep(4)


#Run the first simulation
print ('Run the first simulation')
w_handle = pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
time.sleep(.5)
window.TypeKeys('{F5}')


#Wait on and then dismiss dialog box maximum wait is 20 seconds

w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda: pywinauto.findwindows.find_windows(title= u'Run complete', class_name='#32770')[0])
print ('Dismiss the done dialog box')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


#Start at Top of Table List (GasIN.F)
print ('Start at Top of Table List (GasIN.F)')
w_handle = pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['B1.Inlets Table']
ctrl.SetFocus()
ctrl.ClickInput()
window.TypeKeys('{PGUP}')

#Config Values:
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
time.sleep(.25)
window.TypeKeys('{LEFT}'+'{DOWN}')
time.sleep(.25)
window.TypeKeys(str(GasIN_T)+'{DOWN}')
window.TypeKeys(str(GasIN_zCO2)+'{DOWN}')
window.TypeKeys(str(GasIN_zH2O)+'{DOWN}')
window.TypeKeys(str(GasIN_zN2)+'{DOWN}')
window.TypeKeys(str(SolidIN_Fm)+'{DOWN}')
window.TypeKeys(str(SolidIN_T)+'{DOWN}')
window.TypeKeys(str(SolidIN_wBic)+'{DOWN}')
window.TypeKeys(str(SolidIN_wCar)+'{DOWN}')
window.TypeKeys(str(SolidIN_wH2O)+'{DOWN}')
window.TypeKeys(str(TubeIN_F)+'{DOWN}')
window.TypeKeys(str(TubeIN_P)+'{DOWN}')
window.TypeKeys(str(TubeIN_T)+'{DOWN}')
window.TypeKeys(str(TubeIN_zCO2)+'{DOWN}')
window.TypeKeys(str(TubeIN_zH2O)+'{DOWN}')
window.TypeKeys(str(TubeIN_zN2)+'{DOWN}')


print ('Change Outlets Configure Table')
ctrl = window['B1.Outlets Table']
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
w_handle = pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('{F5}')

#Wait on and then dismiss dialog box, wait 15 seconds maximum.
w_handle = pywinauto.timings.WaitUntilPasses(15,0.5,lambda: pywinauto.findwindows.find_windows(title= u'Run complete', class_name='#32770')[0])
print ('Dismiss the done dialog box')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


print ('Run the Initialization Script')
w_handle = pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.Maximize()
#Force the Icon to the middle of the screen
w_handle = pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE - [Process Flowsheet Window]')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('d')
w_handle = pywinauto.findwindows.find_windows(title=u'Find Object', class_name='#32770')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#Find the icon
ctrl = window['ListBox']
ctrl.Click()
ctrl = window['Find']
ctrl.Click()
window.Close()
w_handle = pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE - [Process Flowsheet Window]')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
# Put mouse on center of window also and click to fill in Flowsheet->Srcipts menu
ctrl.ClickInput()
w_handle = pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE - [Process Flowsheet Window]')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.Restore()
#Start next run using Flowsheet pulldown
window.TypeKeys('%l')
window.TypeKeys('p')
window.TypeKeys('{ENTER}')

#Wait on and then dismiss dialog box, wait up to 45 seconds maximum
w_handle = pywinauto.timings.WaitUntilPasses(45,0.5,lambda: pywinauto.findwindows.find_windows(title= u'Scripting', class_name='#32770')[0])
print ('Dismiss the done dialog box\n\n')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

print ('Test Results Moving Bed Aspen Modeler Simulation '+ Needed_file +' '+ time.asctime())
#Check homotopy correct
#Highlight value to be checked
w_handle = pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['B1.Config Table']
ctrl.SetFocus()
ctrl.ClickInput()
ctrl.TypeKeys('{PGDN}'+'{LEFT}')
ctrl.TypeKeys('{F2}')
ctrl = window['Edit']
ctrl.SetFocus()

#Get expected result

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
   



