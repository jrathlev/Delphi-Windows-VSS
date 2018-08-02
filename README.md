### Delphi Interface to Windows Volume Shadow Service (VSS)

The use of Volume Shadow Copies is described in detail in the Microsoft 
Software Development Kit for Windows 7. As an example, you can find there a 
program (*VSHADOW.EXE*) and the appropriate source code. This, as well as the 
required interfaces (header files) to the system libraries, is however written 
in C++. To use VSS under Delphi, it is first necessary to convert the header 
files into a Delphi unit (file **VssApi.pas**). A second 
unit (**VssUtils.pas**) contains all routines from the Microsoft sample program 
converted to Delphi. To facilitate the integration into user written programs, 
all functions are bundled to a class (**TVolumeShadowCopy**). For execution in an 
own thread, another class (**TVssThread**) is provided. A sample snippet how 
integrate this into a user program can be found in the readme.txt file which 
is part of the source package. 
Additionally, a console application example is provided (**VsToolkit**). This 
application is not based on the original Microsoft SDK sample, but on the modified version 
[Volume Shadow Copy Simple Client (VSCSC)](http://sourceforge.net/projects/vscsc). All 
programs and routines can be compiled for 32- and 64-bit systems (the latter 
requiring at least Delphi XE2).

**Notes:** The routines provided will perform most of the functions needed for backups, 
but a restore is not to date supported.

**Running the program requires administrator rights.**

- Unit **VssApi.pas** - Delphi interface to Windows Volume Shadow Copy Service (*vssapi.dll*)
- Unit **VssUtils.pas** - Delphi objects to use VSS in own programs
- Unit **VssConsts.pas** - language resources for *VssUtils.pas* (English und German)
- Programm **VsToolkit.dpr** - Sample program
