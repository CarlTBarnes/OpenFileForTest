    PROGRAM
!------------------------------------------
! File Lock Check 
! In BAT files before Copy warn if file cannot be locked.
! BAT Example:
!   REM See if FileNameXYZ.TPS is in use, i.e. cannot belocked
!   FileLockCheck FileNameXYZ.TPS
!     IF ERRORLEVEL 2  GOTO :FileAWHOL
!     IF ERRORLEVEL 1  GOTO :FileLocked
!
!------------------------------------------
! History
! 11/17/2009 - New program
!
!------------------------------------------
    include 'tplequ.clw'
    include 'keycodes.clw'
    map
FileLockCheckMain   PROCEDURE()
CanFileBeLocked     PROCEDURE(<*LONG OutErrCode>,<*STRING OutErrText>),BOOL      !returns true if OK
    end

Glo:FileName2Test   STRING(255)
xxxTestFileFlag     BOOL
OpenThisFileName    STRING(260),static
OpenAFileMode       LONG   !AccessModeOption + ShareModeOption
OpenThisFile        FILE,DRIVER('DOS'),NAME(OpenThisFileName),PRE(OpenA)
Record                  RECORD,PRE()
Block                      STRING(1024)   !OpenA:Block
                       END
                    END

    code
    FileLockCheckMain()
    return
!---------------------------------------------
CanFileBeLocked     PROCEDURE(<*LONG OutErrCode>,<*STRING OutErrText>)!,BOOL   !returns true if OK
ErrCode     LONG,AUTO
ErrMsg      STRING(255),AUTO
    CODE
    CLOSE(OpenThisFile)
    OPEN(OpenThisFile,ReadOnly+DenyAll)
    ErrCode = ErrorCode()
    ErrMsg  = Error()
    CLOSE(OpenThisFile)
    IF ErrCode=0 then return 1.
    if ~omitted(OutErrCode) THEN  OutErrCode = ErrCode.
    if ~omitted(OutErrText) THEN  OutErrText = ErrCode & ' ' & ErrMsg.
    return 0
!---------------------------------------------
FileLockCheckMain        procedure
!TextRead        STRING(1024)
!AccessModeOption  BYTE(1)    !default: Read Only
!ShareModeOption  BYTE(1)     !default: Deny All
LastOpenResult   String(255)
LastOpenErrTest  String(255)
LastOpenErrCode  long

WarnedErr2       BOOL
LockCheckWoredLasttime  LONG

Window WINDOW('File Lock Check'),AT(,,296,100),CENTER,GRAY,SYSTEM,ICON('FileLockChk.ico'),FONT('Segoe UI',10),TIMER(300), |
            ALRT(EscKey),DOUBLE
        ENTRY(@s255),AT(5,4,285,17),USE(OpenThisFileName),SKIP,FONT(,10,,FONT:regular),COLOR(COLOR:BTNFACE),READONLY
        BOX,AT(5,26,283,13),USE(?BoxWhite),COLOR(COLOR:White),FILL(COLOR:White),LINEWIDTH(1)
        STRING(@s255),AT(5,28,283,11),USE(LastOpenResult),TRN,CENTER,FONT(,,COLOR:Red,FONT:bold)
        PROMPT('Trying to Lock the above file to insure a clean Copy. Please have other users on the network close any p' & |
                'rogram using this file.'),AT(38,46,219,19),USE(?Prompt1),TRN,CENTER
        BUTTON('&Retry'),AT(77,72,59,17),USE(?RetryBtn)
        BUTTON('&Cancel'),AT(161,72,59,17),USE(?CancelBtn),STD(STD:Close)
        IMAGE('FileLockChk.ICO'),AT(5,65,39,31),USE(?Image1),CENTERED
    END
HaltCode LONG(0)
    code
    Glo:FileName2Test = command('')

    if ~Glo:FileName2Test and exists('testfile.xxx') then
        Glo:FileName2Test= 'TestFile.xxx'
        xxxTestFileFlag=True
    end

    if ~Glo:FileName2Test then
        case message('No File name was specified on the command line' & |
                '||Syntax:     FileLockCheck FILENAME' & |
                '|Help /?' & |
                '||Path: ' & LongPath() &'|EXE:   ' & Command('0'), |
                'FileLockCheck','~FileLockChk.ico','Close|View Help') 
        of 1 ; HALT(2)  
        of 2 ; HaltCode=2 ; Glo:FileName2Test='?'
        end
    end
    
       !Glo:FileName2Test='/?'     !test help
    case lower(Glo:FileName2Test)
    of '?'
    orof '/?'
    orof '\?'
    orof '-?'
    orof 'help'
        message('FileLockCheck tests if a file can be opened for exclusive access.' & |
                '||If the file cannot be opened a window displays telling users to get out. ' & |
                '|This is designed for use in a BAT file to assure a file is closed before COPY. ' & |
                '||Syntax: FileLockCheck FileName' &  |
                '||ErrorLevel set:<9>0=locked|<9,9>1=Could not Lock|<9,9>2=File not found' & |
                '||Batch Example:' & |
                '||<9>FileLockCheck FileNameXYZ.TPS' & |
                '|<9>    IF ERRORLEVEL 2  GOTO :FileAWHOL|<9>    IF ERRORLEVEL 1  GOTO :FileLocked|' & |
                '||Path: ' & LongPath() &'|EXE:   ' & Command('0') & |
                '','FileLockCheck Help','~FileLockChk.ico')
        HALT(HaltCode)
    end

    OpenThisFileName = Glo:FileName2Test    !'D:\Dev5\BCS32\UTIL\WriteIniTest\GetRegAPItest\settings.ini'

    if CanFileBeLocked() THEN   !Can I open this without an Error ?
        halt(0)                 !Yes! I Can so Error Level 0
    end                         !No :( fall thru and open window

    open(window) 
    ?OpenThisFileName{PROP:Background}=80000018h !COLOR:InfoBackground
    ?OpenThisFileName{Prop:FontColor} =80000017h !COLOR:InfoText
    accept
        case event()
        of Event:OpenWindow
            POST(Event:Timer)
        of Event:Timer
           if LockCheckWoredLasttime then halt(0).
           if CanFileBeLocked(LastOpenErrCode,LastOpenErrTest) THEN
              LastOpenResult = format(clock(),@t04)& ' Lock Check completed - File Opened (close in 3 seconds)'
              ?LastOpenResult{PROP:FontColor} = COLOR:Green ; display
              if xxxTestFileFlag then Message('Ready to close','xxxTestFileFlag').
              LockCheckWoredLasttime = 1
              cycle
           end
           LastOpenResult = format(clock(),@t04)& ' Lock Check failed - Error: ' &  LastOpenErrTest
           display
           IF LastOpenErrCode <= 3 and ~WarnedErr2 THEN
              WarnedErr2 = 1
              Message('The ' & choose(LastOpenErrCode=3,'Path','File') & ' name specified cannot be found.' & |
                      '||You should cancel and edit your BAT file.','FileLockCheck Invalid Name')
              select(?CancelBtn)
           END
        end

        case accepted()
        of ?RetryBtn  ; POST(Event:Timer)
        end

    end
    close(window)
    CLOSE(OpenThisFile)
    halt(1)             !User Cancelled
    return
