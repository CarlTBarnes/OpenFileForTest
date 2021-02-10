!Open File for Testing (aka Open A File)
!   As you write code to access files (Open,Copy,Remove) you write 
!   error handling in case that code fails. To be sure that code
!   works as planned you should test it failing. 
!   This tools allows Opening a File in any Access and Share mode 
!   to test your code failing. And test it works with the modes
!   it should like ReadOnly+DenyNone and ReadWrite+DenyNone
!------------------------------------------
! History
! 08/07/2008 - New program, tired of opening files in Word to lock them. 
! 02/07/2021 - Update to Clarion 11. Do2Class routines 
!------------------------------------------

    PROGRAM                            
    INCLUDE 'TPLEQU.CLW'
    MAP
OpenAFile   PROCEDURE()
DB          PROCEDURE(STRING DebugMessage)
FmtHex      PROCEDURE(LONG inHex),STRING
      module('Win32')
        OutputDebugString(*CSTRING cMsg),PASCAL,DLL(1),RAW,NAME('OutputDebugStringA')
        ltoa(LONG Val2Cvt, *CSTRING OutStr, LONG Base),LONG,PROC,RAW,NAME('_ltoa'),DLL(Dll_Mode)
      end
    end
SettingsINI     EQUATE('.\OpenFFT.ini')
    CODE
    OpenAFile()
    RETURN
!---------------------------------------------
OpenAFile        PROCEDURE
OpenThisFileName STRING(260),STATIC
OpenAFileMode    LONG   !AccessModeOption + ShareModeOption
OpenThisFile     FILE,DRIVER('DOS'),NAME(OpenThisFileName),PRE(OpenA)
Record                RECORD,PRE()
Block                    STRING(1024)   !OpenA:Block
                       END
                  END
TextRead          STRING(1024)
AccessModeOption  BYTE(1)    !default: Read Only
ShareModeOption   BYTE(1)    !default: Deny All

Window WINDOW('Open a File for Testing'),AT(,,352,182),GRAY,SYSTEM,ICON('OpenFFT.ICO'),FONT('Segoe UI',10),DROPID('~FILE'), |
            DOUBLE
        PROMPT('&Name:'),AT(4,5),USE(?Prompt1)
        ENTRY(@s255),AT(28,4,284,12),USE(OpenThisFileName)
        BUTTON('&Pick'),AT(319,4,29,12),USE(?PickBtn)
        BUTTON('R&un Again'),AT(298,26,44,15),USE(?ReRunBtn),SKIP,TIP('Run another instance')
        BUTTON('Notepa&d'),AT(298,46,44,15),USE(?NotepadBtn),SKIP,TIP('Run Notepad')
        BUTTON('Explorer'),AT(298,66,44,15),USE(?ExploreBtn),SKIP,TIP('Open Windows Explorer to File')
        OPTION('Access Mode'),AT(27,23,81,49),USE(AccessModeOption),BOXED
            RADIO('Read Only (00h)'),AT(32,33),USE(?AccessModeOption:Radio1)
            RADIO('Write Only (01h)'),AT(32,45),USE(?AccessModeOption:Radio2)
            RADIO('Read/Write (02h)'),AT(32,57),USE(?AccessModeOption:Radio3),TRN
        END
        OPTION('Share Mode'),AT(124,23,81,63),USE(ShareModeOption),BOXED
            RADIO('Deny All (10h)'),AT(129,32),USE(?ShareModeOption:Radio1)
            RADIO('Deny Write (20h)'),AT(129,42),USE(?ShareModeOption:Radio2)
            RADIO('Deny Read (30h)'),AT(129,52),USE(?ShareModeOption:Radio3)
            RADIO('Deny None (40h)'),AT(129,62),USE(?ShareModeOption:Radio4)
            RADIO('Any Access (00h)'),AT(129,72),USE(?ShareModeOption:Radio5)
        END
        BUTTON('&Open File'),AT(225,26,47,15),USE(?OpenFileBtn)
        BUTTON('&Close File'),AT(225,46,47,15),USE(?CloseFileBtn),DISABLE
        BUTTON('&Read File'),AT(225,66,47,15),USE(?ReadFileBtn),DISABLE
        IMAGE,AT(7,76),USE(?Icon:OpenGood)
        IMAGE,AT(37,76),USE(?Icon:OpenFail)
        TEXT,AT(7,95,336,81),USE(TextRead),VSCROLL,FONT('Consolas',11),READONLY
    END
CmdLn   string(260)
CmdX    LONG
Caption PSTRING(64)
DOO     CLASS        !Converted by Do2Class
OpenFileRtn          PROCEDURE()
PickFileRtn          PROCEDURE()
SetOpenIcons         PROCEDURE(BYTE Open1Fail2Hide0) !Opened OK gets Green Check Icon, else Red X 
        END 
    CODE
    OpenThisFileName = '.\TestFiles\OpenTest.Txt'  !Test data
    IF ~EXISTS(OpenThisFileName) THEN
        OpenThisFileName=GETINI('Config','File',,SettingsINI)
    END
    IF command() and exists(command()) then OpenThisFileName = Command().
    TextRead='Test your File code (Open, Copy, Remove) handles errors correctly ' &|
         'when a program has a file locked or opened in an Access or Share ' &|
         'mode not compatible with the mode your code requests. You can also ' &|
         'use this tool to test when your program has a file open what modes ' &|
         'can be used to open and read the file. This tool uses the Clarion ' &|
         'DOS Driver.' & |
             '<13,10,13,10,9>1. Specify your File Name' & |
             '<13,10,9>2. Select Access Mode and Share Mode' & |
             '<13,10,9>3. Press the Open File Button'

    open(window)
    Caption=0{PROP:Text}
    ?OpenThisFileName{PROP:Text}='@s260'    !Runtime picture can exceed 255
    ?TextRead{PROP:Background}=80000018H    !COLOR:InfoBackground EQUATE (80000018H) !Background color for tooltip controls
    ?TextRead{PROP:FontColor} =80000017H    !COLOR:InfoText       EQUATE (80000017H) !Text color for tooltip controls
    DOO.SetOpenIcons(0)
    ?Icon:OpenGood{PROP:Text}=ICON:Tick 
    ?Icon:OpenFail{PROP:Text}=ICON:Cross 
    ?Icon:OpenFail{PROP:XPos}=?Icon:OpenGood{PROP:XPos}

    ACCEPT
        CASE EVENT()
        OF EVENT:Drop
              OpenThisFileName = DROPID()
              DISPLAY
              IF INSTRING(':',OpenThisFileName,1,5) THEN
                 Message('Multiple file names are not supported.||' & DROPID(), Caption )
              END
              IF NOT ?CloseFileBtn{PROP:Disable} THEN POST(EVENT:Accepted, ?CloseFileBtn).
        END
        CASE ACCEPTED()
        of ?PickBtn    ; DOO.PickFileRtn()

        of ?OpenFileBtn
            IF ~OpenThisFileName OR ~EXISTS(OpenThisFileName) THEN
                SELECT(?OpenThisFileName)
                Message(CHOOSE(~OpenThisFileName,'Please enter a file name.','File does not exist.'),Caption)
                CYCLE
            END
            DOO.OpenFileRtn()
            DISPLAY

        of ?CloseFileBtn
            DOO.SetOpenIcons(0)
            CLOSE(OpenThisFile)
            TextRead='File Closed: ' & format(CLOCK(),@t3) &'<13,10><13,10>'&  TextRead
            0{PROP:Text}=Caption
            ENABLE(?OpenFileBtn)
            ENABLE(?OpenThisFileName)
            ENABLE(?PickBtn)
            !ENABLE(?AccessModeOption)
                 ENABLE(?AccessModeOption:Radio1,?AccessModeOption:Radio3)
            !ENABLE(?ShareModeOption)
                 ENABLE(?ShareModeOption:Radio1,?ShareModeOption:Radio5)
            DISABLE(?CloseFileBtn)
            DISABLE(?ReadFileBtn)
            DISPLAY

        of ?ReadFileBtn
           set(OpenThisFile) ; next(OpenThisFile)
           TextRead = OpenA:Block
           DISPLAY

        of ?ReRunBtn ; RUN(command('0'))
        of ?NotepadBtn
            CmdX=1
            IF OpenThisFileName AND EXISTS(OpenThisFileName) THEN
                CmdX=POPUP('Open Notepad|Open File in Notepad')
            END
            IF CmdX THEN
               RUN('Notepad.exe' & CHOOSE(CmdX<>2,'',' ' & CLIP(OpenThisFileName)) )
            END
        of ?ExploreBtn
           IF ~OpenThisFileName OR ~EXISTS(OpenThisFileName) THEN
               RUN('Explorer.exe /e,"' & LongPath() &'"')
           ELSE
               RUN('Explorer.exe /select,"' & CLIP(OpenThisFileName) &'"' )
           END
        end
    end
    close(window)
    close(OpenThisFile)
    PUTINI('Config','File',OpenThisFileName,SettingsINI)
    return
!-------------------------
DOO.OpenFileRtn PROCEDURE()
!     DATA
AccessName  PSTRING(12)
ShareName   PSTRING(12)
AccessLong  LONG
ShareLong   LONG
    CODE
    AccessName=CHOOSE(AccessModeOption,'ReadOnly','WriteOnly','ReadWrite')
    AccessLong=CHOOSE(AccessModeOption, ReadOnly , WriteOnly , ReadWrite )
    ShareName =CHOOSE(ShareModeOption,'DenyAll','DenyWrite','DenyRead','DenyNone','AnyAccess')
    ShareLong =CHOOSE(ShareModeOption, DenyAll , DenyWrite , DenyRead , DenyNone , AnyAccess )
    OpenAFileMode = AccessLong + ShareLong

    TextRead='File ' & CLIP(OpenThisFileName) & |
             '<13,10>' & |
             'Mode: '& AccessName  &' + '& ShareName & |
             ' = ' & FmtHex(AccessLong) &' + '& FmtHex(ShareLong) & |
             ' = ' & FmtHex(OpenAFileMode)

    
    OPEN(OpenThisFile,OpenAFileMode)
    IF ~ERRORCODE() 
       ! 0{PROP:Text}=Caption &': '& CLIP(OpenThisFileName)
        0{PROP:Text}='Open( '& CLIP(OpenThisFileName) &' , '& AccessName  &'+'& ShareName &' )'
        TextRead=clip(TextRead)&'<13,10><13,10>File Opened w/o Error' & |
                                '<13,10>File Status: ' & FmtHex(STATUS(OpenThisFile)) & |
                                '<13,10><13,10>Press the CLOSE Button when you are done to Close the File.'
            DISABLE(?OpenFileBtn)
            DISABLE(?OpenThisFileName)
            DISABLE(?PickBtn)
            !DISABLE(?AccessModeOption)
                 DISABLE(?AccessModeOption:Radio1,?AccessModeOption:Radio3)
                 enable(?AccessModeOption{PROP:ChoiceFEQ})
            !DISABLE(?ShareModeOption)
                 DISABLE(?ShareModeOption:Radio1,?ShareModeOption:Radio5)
                 enable(?ShareModeOption{PROP:ChoiceFEQ})
            ENABLE(?CloseFileBtn)
            ENABLE(?ReadFileBtn)
            DOO.SetOpenIcons(1)
        RETURN
    END
    DOO.SetOpenIcons(2)
    0{PROP:Text}=Caption
    TextRead=clip(TextRead)&'<13,10>' & |
        '<13,10>Error Code: ' & ErrorCode() & ' ' & clip(Error()) & |
        '<13,10>Error File: ' & ErrorFile() & |
        '<13,10>' & choose(Errorcode()<>90,'','File Error: ' & FileErrorcode() & ' ' & clip(FileError())) & |
        ''
    RETURN
!-------------------------
DOO.SetOpenIcons PROCEDURE(BYTE Open1Fail2Hide0) 
HideIco STRING('11')
    CODE 
    IF Open1Fail2Hide0 THEN HideIco[Open1Fail2Hide0]=''.
    ?Icon:OpenGood{PROP:Hide}=HideIco[1]
    ?Icon:OpenFail{PROP:Hide}=HideIco[2] 
    RETURN 
!-------------------------    
DOO.PickFileRtn PROCEDURE()
  CODE
    FILEDIALOG('Select File to Open'     ,|  ! [title]
                   OpenThisFileName      ,|  ! return filename(s)
                                         ,|  ! extensions 'Name|*.ext|Name|*.ext'
                   FILE:LongName          )  ! ? + FILE:KeepDir
    display
    RETURN

!===============================
FmtHex      PROCEDURE(LONG inHex)!,STRING
CFmtd       cstring(64)
    CODE
    ltoa(inHex, CFmtd, 16)
    CFmtd=all('0', 2 - len(CFmtd)) & UPPER(CFmtd) & 'h'
    return  CFmtd

!===============================
DB   PROCEDURE(STRING xMessage)
Prfx EQUATE('OpenFile: ')
sz   CSTRING(SIZE(Prfx)+SIZE(xMessage)+3),AUTO
  CODE 
  sz  = Prfx & CLIP(xMessage) & '<13,10>'
  OutputDebugString( sz )