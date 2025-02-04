# Panzer-Pagefile-Gauge-VB6
 
A FOSS Pagefile Gauge VB6 WoW64 Widget for Windows Vista, 7, 8 and 10/11+. There will also be a version for Reactos and XP, watch this space for the link. Also tested and running well on Linux and Mac os/X using Wine.
 
 My current VB6/RC6 PSD program being finished now, fundamentally complete, only awaiting testing on Windows XP and Win7 32bit and some multi monitor checking, completion of the CHM help file and the creation of the setup.exe. This Panzer widget is based upon the Yahoo widget of the same visual design and very similar operation.

 Why VB6? Well, with a 64 bit, modern-language improvement upgrade on the way with 100% compatible TwinBasic coupled with support for transparent PNGs via RC/Cairo, VB6 code has an amazing future.

![vb6-logo-350](https://github.com/yereverluvinunclebert/Panzer-RAM-Gauge-VB6/assets/2788342/2f60380d-29f5-4737-8392-e7d747c61f25)

 I created as a variation of the previous gauges I had previously created for the World of Tanks and War Thunder 
 communities. The Panzer Pagefile Gauge widget is an attractive dieselpunk VB6 widget for your desktop. 
 Functional and gorgeous at the same time. The graphics are my own, I took original inspiration from a clock face by Italo Fortana combining it with an aircraft gauge surround. It is all my code with some help from the chaps at VBForums (credits given). 
  
The Panzer Pagefile Gauge VB6  is a useful utility displaying the Pagefile usage of your system in a dieselpunk fashion on your desktop. This Widget is a moveable widget that you can move anywhere around the desktop as you require.

The following is the code used to extract the memory information from the running system via the GlobalMemoryStatusEx API. These are the pertinent bits:


    Private Declare Sub GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX)
    
    Private Type INT64
       LoPart As Long
       HiPart As Long
    End Type

    Private Type MEMORYSTATUSEX
       dwLength As Long
       dwMemoryLoad As Long
       ulTotalPhys As INT64
       ulAvailPhys As INT64
       ulTotalPageFile As INT64
       ulAvailPageFile As INT64
       ulTotalVirtual As INT64
       ulAvailVirtual As INT64
       ulAvailExtendedVirtual As INT64
    End Type    
    
    Private Function Pagefile_Usage_Percent() As Long
        Dim udtMemStatEx As MEMORYSTATUSEX
        Dim physMemFree As Long: physMemFree = 0
    
        udtMemStatEx.dwLength = Len(udtMemStatEx)
        Call GlobalMemoryStatusEx(udtMemStatEx)
    
        physMemFree = Round(CLargeInt(udtMemStatEx.ulAvailPageFile.LoPart, udtMemStatEx.ulAvailPageFile.HiPart) / (CLargeInt(udtMemStatEx.ulTotalPageFile.LoPart, udtMemStatEx.ulTotalPageFile.HiPart)) * 100)
    
        Pagefile_Usage_Percent = 100 - physMemFree
    End Function

    Private Function Pagefile_Usage_Total() As Double
        Dim udtMemStatEx As MEMORYSTATUSEX
        Dim physMem As Double: physMem = 0
    
        udtMemStatEx.dwLength = Len(udtMemStatEx)
        all GlobalMemoryStatusEx(udtMemStatEx)

        physMem = NumberInKB(CLargeInt(udtMemStatEx.ulTotalPageFile.LoPart, udtMemStatEx.ulTotalPageFile.HiPart))
    
        Pagefile_Usage_Total = physMem
    End Function
    
    Private Function obtainPagefileDetails() As String
      Dim result As String: result = vbNullString
      Dim udtMemStatEx As MEMORYSTATUSEX
      
      udtMemStatEx.dwLength = Len(udtMemStatEx)
      Call GlobalMemoryStatusEx(udtMemStatEx)
      
      On Error GoTo obtainPagefileDetails_Error
  
      result = result & "System Pagefile Total: " & Trim$(Pagefile_Usage_Total) & sizString & "  , " & vbCrLf
      result = result & "Available Pagefile: " + vbTab + NumberInKB2(CLargeInt(udtMemStatEx.ulAvailPageFile.LoPart, udtMemStatEx.ulAvailPageFile.HiPart)) & vbCrLf
      result = result & "Percentage Usage: " + vbTab + CStr(Pagefile_Usage_Percent) & "% utilised." & vbCrLf
      result = result & "Available extended page file " + vbTab + NumberInKB2(CLargeInt(udtMemStatEx.ulAvailExtendedVirtual.LoPart, udtMemStatEx.ulAvailExtendedVirtual.HiPart)) & vbCrLf
     
      Debug.Print result
  
      obtainPagefileDetails = result
    End Function

There are some other bits of code used for formatting the results that you will need. Just dig into the code of this desktop gauge to find the above and it will become apparent which other bits you require. 
 
![pagefilePhoto](https://github.com/user-attachments/assets/07750de4-32d0-4810-a22a-e1a480cb7a6b)

 This widget can be increased in size, animation speed can be changed, 
 opacity/transparency may be set as to the users discretion. The widget can 
 also be made to hide for a pre-determined period.
 
![pagefileGauge](https://github.com/user-attachments/assets/c1e404b0-883c-4ffe-80e7-22ac60ae8f93)

 Right clicking will bring up a menu of options. Double-clicking on the widget will cause a personalised Windows application to 
 fire up. The first time you run it there will be no assigned function and so it 
 will state as such and then pop up the preferences so that you can enter the 
 command of your choice. The widget takes command line-style commands for 
 windows. Mouse hover over the widget and press CTRL+mousewheel up/down to resize. It works well on Windows XP 
 to Windows 11.
 
![panzergauge-help](https://github.com/user-attachments/assets/f9bb0e1d-9fa3-4d8d-9f42-d52de15b649a)

 The Panzer Pagefile Gauge VB6 gauge is Beta-grade software, under development, not yet 
 ready to use on a production system - use at your own risk.

 This version was developed on Windows 7 using 32 bit VisualBasic 6 as a FOSS 
 project creating a WoW64 widget for the desktop. 
 
![Licence002](https://github.com/yereverluvinunclebert/Panzer-RAM-Gauge-VB6/assets/2788342/09dd88fd-0bff-4115-8fda-9b3e6b6852f5)


 It is open source to allow easy configuration, bug-fixing, enhancement and 
 community contribution towards free-and-useful VB6 utilities that can be created
 by anyone. The first step was the creation of this template program to form the 
 basis for the conversion of other desktop utilities or widgets. A future step 
 is new VB6 widgets with more functionality and then hopefully, conversion of 
 each to RADBasic/TwinBasic for future-proofing and 64bit-ness. 
 
![menu01](https://github.com/yereverluvinunclebert/Panzer-RAM-Gauge-VB6/assets/2788342/ee727437-e6e4-4b91-8c0d-90e7e43352b4)

 This utility is one of a set of steampunk and dieselpunk widgets. That you can 
 find here on Deviantart: https://www.deviantart.com/yereverluvinuncleber/gallery
 
 I do hope you enjoy using this utility and others. Your own software 
 enhancements and contributions will be gratefully received if you choose to 
 contribute.

 BUILD: The program runs without any Microsoft plugins.
 
 Built using: VB6, MZ-TOOLS 3.0, VBAdvance, CodeHelp Core IDE Extender
 Framework 2.2 & Rubberduck 2.4.1, RichClient 6
 
 Links:
 
	https://www.vbrichclient.com/#/en/About/
	MZ-TOOLS https://www.mztools.com/  
	CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1  
	Rubberduck http://rubberduckvba.com/  
	Rocketdock https://punklabs.com/  
	Registry code ALLAPI.COM  
	La Volpe http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1  
	PrivateExtractIcons code http://www.activevb.de/rubriken/  
	Persistent debug code http://www.vbforums.com/member.php?234143-Elroy  
	Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain  
	VBAdvance  
 
![pagefileBalloon](https://github.com/user-attachments/assets/d48b51eb-96fd-4f82-9188-bfa0a67fe0fd)
 
 Tested on :
 
	ReactOS 0.4.14 32bit on virtualBox    
	Windows 7 Professional 32bit on Intel    
	Windows 7 Ultimate 64bit on Intel    
	Windows 7 Professional 64bit on Intel    
	Windows XP SP3 32bit on Intel    
	Windows 10 Home 64bit on Intel    
	Windows 10 Home 64bit on AMD    
	Windows 11 64bit on Intel  

	
   
 CREDITS:
 
 I have really tried to maintain the credits as the project has progressed. If I 
 have made a mistake and left someone out then do forgive me. I will make amends 
 if anyone points out my mistake in leaving someone out.
 
 MicroSoft in the 90s - MS built good, lean and useful tools in the late 90s and 
 early 2000s. Thanks for VB6.
 
 Olaf Schmidt - This tool was built using the RichClient RC5 Cairo wrapper for 
 VB6. Specifically the components using transparency and reading images directly 
 from PSD. Thanks for the massive effort Olaf in creating Cairo counterparts for 
 all VB6 native controls and giving us access to advanced features on controls 
 such as transparency.
 
 Shuja Ali @ codeguru for his settings.ini code.
 
 ALLAPI.COM        For the registry reading code.
 
 Rxbagain on codeguru for his Open File common dialog code without a dependent 
 OCX - http://forums.codeguru.com/member.php?92278-rxbagain
 
 si_the_geek       for his special folder code
 
 Elroy on VB forums for the balloon tooltips

 Hypetia for her physMem code - https://www.tek-tips.com/userinfo.cfm?member=Hypetia

 Khodor for the RAM numbers conversion code https://stackoverflow.com/users/164214/khodor
 
 Harry Whitfield for his quality testing, brain stimulation and being an 
 unwitting source of inspiration. 
 
 Dependencies:
 
 o A windows-alike o/s such as Windows XP, 7-11 or Apple Mac OSX 11. 
 
 o Microsoft VB6 IDE installed with its runtime components. The program runs 
 without any additional Microsoft OCX components, just the basic controls that 
 ship with VB6.  
 
![vb6-logo-350](https://github.com/yereverluvinunclebert/Panzer-RAM-Gauge-VB6/assets/2788342/2479af5a-82bf-42ae-bdb1-28c22160f93c)
 
 * Uses the latest version of the RC6 Cairo framework from Olaf Schmidt.
 
 During development the RC6 components need to be registered. These scripts are 
 used to register. Run each by double-clicking on them.
 
	RegisterRC6inPlace.vbs
	RegisterRC6WidgetsInPlace.vbs
 
 During runtime on the users system, the RC6 components are dynamically 
 referenced using modRC6regfree.bas which is compiled into the binary.	
 
 
 Requires a PzPagefile Gauge folder in C:\Users\<user>\AppData\Roaming\ 
 eg: C:\Users\<user>\AppData\Roaming\PzPagefile Gauge
 Requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\PzPagefile Gauge
 The above will be created automatically by the compiled program when run for the 
 first time.
 
 
o Krool's replacement for the Microsoft Windows Common Controls found in
mscomctl.ocx (slider) are replicated by the addition of one
dedicated OCX file that are shipped with this package.

During development only, this must be copied to C:\windows\syswow64 and should be registered.

- CCRSlider.ocx

Register this using regsvr32, ie. in a CMD window with administrator privileges.
	
	c:                          ! set device to boot drive with Windows
	cd \windows\syswow64s	    ! change default folder to syswow64
	regsvr32 CCRSlider.ocx	! register the ocx

This will allow the custom controls to be accessible to the VB6 IDE
at design time and the sliders will function as intended (if this ocx is
not registered correctly then the relevant controls will be replaced by picture boxes).

The above is only for development, for ordinary users, during runtime there is no need to do the above. The OCX will reside in the program folder. The program reference to this OCX is contained within the supplied resource file, Panzer Pagefile Gauge.RES. The reference to this file is already compiled into the binary. As long as the OCX is in the same folder as the binary the program will run without the need to register the OCX manually.
 
 * OLEGuids.tlb
 
 This is a type library that defines types, object interfaces, and more specific 
 API definitions needed for COM interop / marshalling. It is only used at design 
 time (IDE). This is a Krool-modified version of the original .tlb from the 
 vbaccelerator website. The .tlb is compiled into the executable.
 For the compiled .exe this is NOT a dependency, only during design time.
 
 From the command line, copy the tlb to a central location (system32 or wow64 
 folder) and register it.
 
 COPY OLEGUIDS.TLB %SystemRoot%\System32\
 REGTLIB %SystemRoot%\System32\OLEGUIDS.TLB
 
 In the VB6 IDE - project - references - browse - select the OLEGuids.tlb
 
![prefs-about](https://github.com/user-attachments/assets/853f9d1d-cb69-4bfc-92d0-61d07e279b7e)

 * SETUP.EXE - The program is currently distributed using setup2go, a very useful 
 and comprehensive installer program that builds a .exe installer. Youll have to 
 find a copy of setup2go on the web as it is now abandonware. Contact me
 directly for a copy. The file "install PzRAM Gauge 0.1.0.s2g" is the configuration 
 file for setup2go. When you build it will report any errors in the build.
 
 * HELP.CHM - the program documentation is built using the NVU HTML editor and 
 compiled using the Microsoft supplied CHM builder tools (HTMLHelp Workshop) and 
 the HTM2CHM tool from Yaroslav Kirillov. Both are abandonware but still do
 the job admirably. The HTML files exist alongside the compiled CHM file in the 
 HELP folder.
 
  Project References:

	VisualBasic for Applications  
	VisualBasic Runtime Objects and Procedures  
	VisualBasic Objects and Procedures  
	OLE Automation  
	vbRichClient6  
 
 
 LICENCE AGREEMENTS:
 
 Copyright © 2023 Dean Beedell
 
 In addition to the GNU General Public Licence please be aware that you may use 
 any of my own imagery in your own creations but commercially only with my 
 permission. In all other non-commercial cases I require a credit to the 
 original artist using my name or one of my pseudonyms and a link to my site. 
 With regard to the commercial use of incorporated images, permission and a 
 licence would need to be obtained from the original owner and creator, ie. me.
 
![about](https://github.com/user-attachments/assets/b445d44f-5b56-47d4-9447-988ef36444a1)


![Panzer-CPU-Gauge-onDesktop](https://github.com/yereverluvinunclebert/Panzer-RAM-Gauge-VB6/assets/2788342/6dc97b14-9954-4f8c-a775-5657b2aeec85)





 
