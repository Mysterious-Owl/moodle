import PySimpleGUI as sg
import os
import sys
from zipfile import ZipFile
import requests
import win32com.client

# pyinstaller --uac-admin --onefile -wF Installer_MoodleProgram.py

timetable_data = '''[Monday]
start=11
end=17
11=DS
12=NPTEL
13=LUNCH
14=Optical
15=Microprocessor Lab
16=Microprocessor Lab

[Tuesday]
start=10
end=17
10=DM
11=DSP
12=Optical
13=LUNCH
14=
15=
16=NPTEL

[Wednesday]
start=11
end=17
11=DS
12=DSP
13=LUNCH
14=Minor Project Lab
15=Minor Project Lab
16=MandI

[Thursday]
start=10
end=16
10=DM
11=MandI
12=Optical
13=LUNCH
14=Minor Project Lab
15=Minor Project Lab

[Friday]
start=11
end=16
11=DS
12=
13=LUNCH
14=DSP
15=MandI'''

config_data = '''[login0]
user={}
pass={}

[links]
ds = http://moodle.mitsgwalior.in/mod/attendance/view.php?id=48597
optical = http://moodle.mitsgwalior.in/mod/attendance/view.php?id=49314
microprocessor lab = http://moodle.mitsgwalior.in/mod/attendance/view.php?id=48591
dm = http://moodle.mitsgwalior.in/mod/attendance/view.php?id=7142
dsp = http://moodle.mitsgwalior.in/mod/attendance/view.php?id=48692
minor project lab = http://moodle.mitsgwalior.in/mod/attendance/view.php?id=49726
mandi = http://moodle.mitsgwalior.in/mod/attendance/view.php?id=48916'''

sg.theme('Dark')

ins_path = ''
for i in os.environ['USERPROFILE']:
    if i == '\\':
        ins_path += '/'
    else:
        ins_path += i

event, values = sg.Window('Installer',
                          [[sg.Text('', size=(2, 3))],
                           [sg.Text('Select installation directory')],
                           [sg.In(default_text=ins_path),
                            sg.FolderBrowse(initial_folder=ins_path)],
                           [sg.Text('Will create a new folder there: Mysterious Owl')],
                           [sg.Text('', size=(2, 3))],
                           [sg.Text('', size=(39, 1)), sg.Button('Next', key='next', size=(6, 1))]],
                          finalize=True).read(close=True)
if event == sg.WIN_CLOSED:
    sys.exit(0)
fname = values[0]
if fname.endswith('/'):
    fname = fname[:-1]
ins_path = ''

for i in fname:
    if i == '/':
        ins_path += '\\'
    else:
        ins_path += i

ins_path += '\\Mysterious Owl\\'
moodle_path = ins_path + 'moodleprogram\\'
config_path = moodle_path + 'config\\'

event, values = sg.Window('Installer',
                          [[sg.Text('Enter login details -', size=(20, 2), text_color='yellow')],
                           [sg.Text('Username:', size=(8, 1)), sg.Input(key='user')],
                           [sg.Text('Password:', size=(8, 1)), sg.Input(key='pass')],
                           [sg.Text('', size=(2, 1))],
                           [sg.Checkbox('Create Desktop shortcut', key='desk')],
                           [sg.Checkbox('Create Start shortcut', key='star')],
                           [sg.Checkbox('Keep all old configuration files?', key='con')],
                           [sg.Text('', size=(2, 1))],
                           [sg.Text('', size=(41, 1)), sg.Button('Next', key='next', size=(6, 1))]],
                          finalize=True).read(close=True)

if event == sg.WIN_CLOSED:
    sys.exit(0)
user = values['user']
password = values['pass']
desk = values['desk']
star = values['star']
con = values['con']

pro_win = sg.Window('Installer',
                    [[sg.Text('Progress:', size=(20, 2), text_color='yellow')],
                     [sg.Text(size=(50, 6), key='prokey')],
                     [sg.Text('', size=(2, 1))],
                     [sg.Text('Installation Complete!!', size=(20, 2), text_color='yellow', visible=False, key="comp")],
                     [sg.Text('', size=(2, 1))],
                     [sg.Text('', size=(20, 1)), sg.Button('Launch', key='launch', size=(6, 1), visible=False),
                      sg.Button('Exit', key='exit', size=(6, 1), visible=False)]
                     ], finalize=True)

try:
    os.mkdir(ins_path)
except FileExistsError:
    pass

if con:
    if os.path.exists(config_path + "user.ini"):
        with open(config_path + "user.ini", "r") as f:
            config_data = f.read()
    if os.path.exists(config_path + "timetable.ini"):
        with open(config_path + "timetable.ini", "r") as f:
            timetable_data = f.read()

progress = ''
progress += 'Downloading Files..\n'
pro_win['prokey'].update(progress)
button, values = pro_win.read(timeout=10)

open(ins_path + 'moodleprogram.zip', 'wb').write(
    requests.get("https://github.com/Mysterious-Owl/moodle/releases/latest/download/moodleprogram.zip",
                 allow_redirects=True).content)

open(ins_path + 'chromedriver.zip', 'wb').write(requests.get("https://chromedriver.storage.googleapis.com/91.0.4472.19"
                                                             "/chromedriver_win32.zip", allow_redirects=True).content)

progress += 'Extracting Files..\n'
pro_win['prokey'].update(progress)
button, values = pro_win.read(timeout=10)

with ZipFile(ins_path + 'moodleprogram.zip', 'r') as f1:
    f1.extractall(ins_path)

with ZipFile(ins_path + 'chromedriver.zip', 'r') as f2:
    f2.extractall(config_path)

progress += 'Deleting Obsolete Files..\n'
pro_win['prokey'].update(progress)
button, values = pro_win.read(timeout=10)

os.remove(ins_path + "moodleprogram.zip")
os.remove(ins_path + "chromedriver.zip")

progress += 'Creating Files..\n'
pro_win['prokey'].update(progress)
button, values = pro_win.read(timeout=10)

try:
    os.mkdir(ins_path + '\\batch')
except FileExistsError:
    pass

with open(ins_path + "batch\\moodle.bat", "w") as f:
    f.write('''@echo off
SET mypath=%~dp0
pushd %mypath%
cd ..
cd moodleprogram
moodleprogram.exe %*
popd
timeout 2''')

with open(config_path + "timetable.ini", "w") as f:
    f.write(timetable_data)

with open(config_path + "user.ini", "w") as f:
    f.write(config_data.format(user, password))

with open(ins_path + "\\moodle.xml", "w") as f:
    f.write('''<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2021-01-21T19:50:58.321446</Date>
    <Author>MYSTERIOUS-OWL</Author>
    <Description>To mark automatic attendance</Description>
    <URI>\\moodle</URI>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>2021-01-21T10:28:00</StartBoundary>
      <ExecutionTimeLimit>PT30M</ExecutionTimeLimit>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Monday />
          <Tuesday />
          <Wednesday />
          <Thursday />
          <Friday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
    <CalendarTrigger>
      <StartBoundary>2021-01-21T11:28:00</StartBoundary>
      <ExecutionTimeLimit>PT30M</ExecutionTimeLimit>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Monday />
          <Tuesday />
          <Wednesday />
          <Thursday />
          <Friday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
    <CalendarTrigger>
      <StartBoundary>2021-01-21T12:28:00</StartBoundary>
      <ExecutionTimeLimit>PT30M</ExecutionTimeLimit>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Monday />
          <Tuesday />
          <Wednesday />
          <Thursday />
          <Friday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
    <CalendarTrigger>
      <StartBoundary>2021-01-21T14:28:00</StartBoundary>
      <ExecutionTimeLimit>PT30M</ExecutionTimeLimit>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Monday />
          <Tuesday />
          <Wednesday />
          <Thursday />
          <Friday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
    <CalendarTrigger>
      <StartBoundary>2021-01-21T15:28:00</StartBoundary>
      <ExecutionTimeLimit>PT30M</ExecutionTimeLimit>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Monday />
          <Tuesday />
          <Wednesday />
          <Thursday />
          <Friday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
    <CalendarTrigger>
      <StartBoundary>2021-01-21T16:28:00</StartBoundary>
      <ExecutionTimeLimit>PT30M</ExecutionTimeLimit>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Monday />
          <Tuesday />
          <Wednesday />
          <Thursday />
          <Friday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <GroupId>S-1-5-32-545</GroupId>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT15M</Interval>
      <Count>2</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>''' + ins_path + 'batch\\' + '''moodle.bat</Command>
      <Arguments>auto</Arguments>
      <WorkingDirectory>''' + ins_path + 'batch\\' + '''</WorkingDirectory>
    </Exec>
  </Actions>
</Task>   
''')

progress += 'Creating ShortCuts..\n'
pro_win['prokey'].update(progress)
button, values = pro_win.read(timeout=10)

desktop = r'{}'.format(ins_path)
path = os.path.join(desktop, 'moodleprogram.lnk')
target = r'{}moodleprogram.exe'.format(moodle_path)
icon = r'{}icon\favicon.ico'.format(moodle_path)

shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(path)
shortcut.Targetpath = target
shortcut.WorkingDirectory = r'{}'.format(moodle_path)
shortcut.IconLocation = icon
shortcut.WindowStyle = 1
shortcut.save()

shor_path = ins_path + "\\moodleprogram.lnk"

if desk:
    os.system("copy \"" + shor_path + "\" \"" + os.environ['HOMEDRIVE'] + os.environ['HOMEPATH'] + "\\Desktop\"")
if star:
    try:
        os.mkdir(os.environ['APPDATA'] + "\\Microsoft\\Windows\\Start Menu\\Programs\\Mysterious Owl")
    except FileExistsError:
        pass
    os.system("copy \"" + shor_path + "\" \"" + os.environ[
        'APPDATA'] + "\\Microsoft\\Windows\\Start Menu\\Programs\\Mysterious Owl\"")

Aaa = '''
1. To automate attendance goto 
   "Start -> Windows Administrative Tools -> Task Scheduler" or search "Task Scheduler".

2. Click on "Task Scheduler Library" on right side then on left side click on "Import Task".

3. Browse to " ''' + ins_path + ''' " and select "moodle.xml" and then OK then OK.

4. Done
'''
sg.popup(Aaa, title="Automate Attendance", line_width=100, custom_text="Got it!", font=("", 12), text_color='sky blue',
         background_color='black')
pro_win['comp'].update(visible=True)
pro_win['launch'].update(visible=True)
pro_win['exit'].update(visible=True)

while 1:
    event, value = pro_win.read()
    if event == 'launch':
        Flag = True
        break
    elif event == 'exit' or event == sg.WIN_CLOSED or event == 'Exit':
        pro_win.Close()
        Flag = False
        break
if Flag:
    pro_win.Close()
    os.system("\"" + shor_path + "\"")
