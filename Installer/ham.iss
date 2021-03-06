; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
AppName=Hours and Minutes
AppVerName=Hours and Minutes 1.6
AppPublisher=Port Jackson Computing
AppPublisherURL=http://www.hoursandminutes.com
AppSupportURL=http://www.hoursandminutes.com
AppUpdatesURL=http://www.hoursandminutes.com
DefaultDirName={pf}\Hours and Minutes
DefaultGroupName=Hours and Minutes
AllowNoIcons=yes
LicenseFile=..\Hours and Minutes License.rtf
InfoAfterFile=..\Hours and Minutes Readme.rtf
; uncomment the following line if you want your installation to run on NT 3.51 too.
; MinVersion=4,3.51

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; MinVersion: 4,4

[Files]
Source: "files\application folder\*.*"; DestDir: "{app}"; CopyMode: alwaysoverwrite

Source: "files\system32\asycfilt.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall
Source: "files\system32\comcat.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall
Source: "files\system32\comdlg32.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall
Source: "files\system32\hhctrl.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall
Source: "files\system32\itircl.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall
Source: "files\system32\itss.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall
Source: "files\system32\mscomct2.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall
Source: "files\system32\mscomctl.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall
Source: "files\system32\msvbvm60.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall
Source: "files\system32\oleaut32.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall
Source: "files\system32\olepro32.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall
Source: "files\system32\stdole2.tlb"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall
Source: "files\system32\VB6STKIT.DLL"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: uninsneveruninstall


[Icons]
Name: "{group}\Hours and Minutes"; Filename: "{app}\Hours and Minutes.exe"
Name: "{group}\Hours and Minutes"; Filename: "{app}\Hours and Minutes Help.chm"

[Run]
Filename: "{app}\Hours and Minutes.exe"; Description: "Launch Hours and Minutes"; Flags: nowait postinstall skipifsilent

