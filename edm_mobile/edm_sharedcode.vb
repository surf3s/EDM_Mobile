Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Imports System
Imports System.IO.Ports
Imports Microsoft.VisualBasic

' To do next
'   positioning of east/west north/south up/down in form
'   need view readme form/file
'   viewing db grids needs an hourglass
'   update of grid in points does not workx
'   saving/displaying time apparently does not work

' After visit in June 2011
'   make sure the so and so problem of ID numbers is fixed
'   program doesn't handle well a db path that doesn't exist
'   program doesn't handle well a new field added to cfg
'   move barcode code to all other text fields
'   Unit fields isn't right
'       what the program does is store everything there and recall everything
'       ie. it is ignoring what has been called a unit field - check relation to carry as well
'   Need to move vartext1 code to the other vartext fields

Module edm_sharedcode

    Public commcontrol As New SerialPort            'the COM port shared between all forms

    Public WriteLogFile As Boolean                  'determines if a log file is written
    Public showoffsetmenu As Boolean                'determines whether prism constants show in menus
    Public donotwaitforprism As Boolean = False     'sets whether to wait for button press after shot to accept prism
    Public measurementadjustment As Double          'passes value from offset form to editpoint form
    Public menuresult As String                     'passes value from menulist back to editpoint
    Public returnedstatus As String                 'passes a status message back to main

    'Used by the plot routine
    Public overlayfile As String = ""

    Public tablesopen As Boolean

    Public fpath As String
    Public cfgfile As String
    Public edmini As String
    Public LogFile As String

    Public bar_code_unit_id As String = ""
    Public bar_code_pre_scan As Boolean = False

    Public button_values(6, 100) As String

    Public limits(100, 4) As Double
    Public limitname(100) As String
    Public unitlimits As Integer
    Public use_limitchecking As Boolean
    Public unitname As String

    Public datum(100, 4) As Double
    Public datumname(100) As String
    Public ndatums As Integer

    Public npoles As Integer
    Public poledata(100, 2) As Single
    Public polename(100) As String

    'These variables handle the results of the last shot
    Public currentstationx As Double
    Public currentstationy As Double
    Public currentstationz As Double
    Public currentstationname As String
    Public referencedatum As String
    Public referencedatum2 As String
    Public referencedatum3 As String
    Public verifydatum As String

    Public firstxp As Double
    Public firstyp As Double
    Public firstzp As Double

    Public currentprismheight As Single
    Public currentprismoffset As Single
    Public currentprismname As String
    Public edmprismoffset As Single
    Public currentxp As Double
    Public currentyp As Double
    Public currentzp As Double
    Public currentsloped As Double
    Public currentvangle As String
    Public currenthangle As String
    Public shottype As Integer      '0=no shot, >0 record no. of shot being edited, -1=new shot, -2=x-shot
    Public button_no As Integer     'for user defined buttons
    Public in_emulation As Boolean  '
    Public back_to_edit As Boolean
    Public back_to_main As Boolean
    Public shot_canceled As Boolean
    Public currentunit As String
    Public refresh_point_grid As Boolean
    Public last_id As String
    Public last_unit As String
    Public last_suffix As String

    Public editvar As Integer       'which variable is being edited at any one moment
    Public actualvar As Integer     'which variable in the list vs. editvar which is the control
    Public responses(100) As String 'used to store the values for a point while they are edited

    Public vars As Integer

    Public vtype(100) As Integer
    Public vardefault(100) As String
    Public varprompt(100) As String
    Public varlist(100) As String
    Public vlen(100) As Integer
    Public vmenu(100) As String
    Public vunique(100) As Boolean
    Public vcarry(100) As Boolean
    Public vincr(100) As Boolean
    Public varprint(100) As Boolean
    Public varrequired(100) As Boolean

    Public options(100) As String
    'option 1 - output filename (mdb)
    'option 2 - com parameters
    'option 3 - which com
    'option 4 - edm type
    'option 5
    'option 6
    'option 7
    'option 8
    'option 9
    'option 10
    'option 11 - computer type
    'option 12 - sqid system yes or no
    'option 13 - pole names
    'option 14 - pole heights
    'option 15 - pole offset
    'option 16
    'option 17 - unit fields
    'option 18 - site name
    'option 19 - printer yes or no
    'option 20
    'option 21 - use vhd or xyz for manual input
    'option 22 - output table name

    Sub parsecfg(ByRef needtoconvert As Integer)

        '--------------------------------------------------
        ' Open the config file passed from main module
        ' Note that we already know it does in fact exist.
        '--------------------------------------------------

        Dim lineno As Integer
        Dim fl As String
        Dim ts As String
        Dim a As Integer
        Dim b As Integer
        Dim c As Integer
        Dim d As Integer
        Dim e As Integer
        Dim f As Integer
        Dim key As String
        Dim dt As String
        Dim comport As String = ""
        Dim temp(100) As String
        Dim edmce As Boolean
        Dim answer As Integer
        Dim tempunitname As String
        Dim varn As String
        Dim dat As String
        Dim errorcode As Integer
        Dim errorflag As Integer
        Dim errormessage As String

        Dim inidata(100, 2) As String
        Dim iniclass As String
        Dim status As Integer
        Dim l As String

        npoles = 0
        vars = 0
        ndatums = 0
        needtoconvert = 0

        'first, check to see if the file is in old EDM format
        'or new EDMCE/EDM format by looking for that section in the file
        edmce = False
        Dim file1 As New System.IO.StreamReader(cfgfile)
        l = file1.ReadLine
        Do Until l Is Nothing
            If InStr(UCase(l), "[EDM]") <> 0 Then
                edmce = True
                Exit Do
            End If
            l = file1.ReadLine
        Loop
        file1.Close()
        file1.Dispose()

        If Not edmce Then
            lineno = 0
            Dim file2 As New System.IO.StreamReader(cfgfile)
            Do
                '--------------------------------------------------
                ' Look for a continuation character at the end of
                ' the line.  If it exists add the next line to a.
                ' In the end, a contains a full line from the CFG.
                '--------------------------------------------------
                fl = ""
                Do
                    lineno = lineno + 1
                    ts = file2.ReadLine
                    If ts Is Nothing Then Exit Do
                    ts = Trim(ts)
                    fl = fl + ts
                    If Right(ts, 1) = "\" Then Mid(fl, Len(fl), 1) = " "
                Loop While Right(ts, 1) = "\"

                '-------------------------------------------
                ' Check for an empty line or for a ' comment
                '-------------------------------------------
                If (Len(fl) <> 0 Or fl <> Space(Len(fl))) And Left(fl, 1) <> "'" Then

                    '----------------------------------------------
                    ' Look for : as a marker for keywords
                    '----------------------------------------------
                    a = InStr(fl, ":")
                    If a <> 0 Then

                        '--------------------------------------------
                        ' Set key to the keyword then select case on
                        ' the different types of keyword.
                        '--------------------------------------------
                        key = UCase(Trim(Left(fl, a - 1)))
                        dt = Trim(Mid(fl, a + 1))
                        Select Case key
                            Case "FIELD"

                                If dt = "" Then
                                    errorcode = 10
                                    errormessage = "No field name given. "
                                    Exit Do
                                End If

                                '-------------------------------------
                                ' Check against the array of all field
                                ' names for duplicates.
                                '-------------------------------------
                                For b = 1 To vars
                                    If varlist(b) = dt Then
                                        errorcode = 1
                                        errormessage = "Duplicate FIELD definition for " + dt
                                        Exit Do
                                    End If
                                Next b

                                '-------------------------------------
                                ' Check the upper limit for number of
                                ' possible variables in the program.
                                '-------------------------------------
                                If vars = UBound(varlist, 1) Then
                                    errorcode = 12
                                    errormessage = "Too many FIELD definitions.  Limit is" + Str(UBound(varlist, 1))
                                    Exit Do
                                End If

                                '-------------------------------------
                                ' If a new field name add this to the
                                ' field name list.
                                '-------------------------------------
                                vars = vars + 1
                                varlist(vars) = UCase(dt)

                            Case "POINTSFILE"
                                '-------------------------------------
                                'Set option(1) to the output filename
                                '-------------------------------------
                                options(1) = UCase(dt)

                            Case "COM1"
                                '-------------------------------------
                                'Set option(2) to the com parameters
                                '-------------------------------------
                                options(2) = "COM1"
                                options(3) = UCase(dt)

                            Case "COM2"
                                '-------------------------------------
                                'Set option(3) to the com parameters
                                '-------------------------------------
                                options(2) = "COM2"
                                options(3) = UCase(dt)

                            Case "EDM"
                                '-------------------------------------
                                'Set option(4) to the type of instrument
                                'and check that it is a valid type.
                                '-------------------------------------
                                ts = UCase(dt)
                                d = InStr(ts, " ")
                                If d = 0 And ts <> "NONE" Then
                                    errorcode = 11
                                    errormessage = "Missing COM port assignment."
                                    Exit Do
                                ElseIf ts <> "NONE" Then
                                    comport = Trim(Mid(ts, d))
                                    options(4) = UCase(LTrim(Left(ts, d - 1)))
                                Else
                                    options(4) = "NONE"
                                End If
                                Select Case options(4)
                                    Case "GTS-3B", "GTS-301", "GTS-302", "GTS-303", "GTS-3X", "TOPCON"
                                        options(4) = "TOPCON" + " " + comport
                                    Case "NONE"
                                    Case "MANUAL"
                                    Case "TC-500", "WILD"
                                        options(4) = "WILD" + " " + comport
                                    Case "SOKKIA"
                                        options(4) = "SOKKIA" + " " + comport
                                    Case Else
                                        errorcode = 11
                                        errormessage = "Expect TOPCON, WILD or SOKKIA instrument."
                                        Exit Do
                                End Select

                            Case "PRINTFIELDS"

                            Case "TEXTFOR"

                            Case "TEXTBACK"

                            Case "SQIDFILE"

                                '--------------------------------------------------
                                'Option to know computer type (ie. HP vs PC)
                                '--------------------------------------------------
                            Case "COMPUTER"
                                options(11) = UCase(dt)
                                Select Case options(11)
                                    Case "HP"
                                        options(11) = "HP"
                                    Case "HP-95", "HP-95LX", "HP95"
                                        options(11) = "HP95"
                                    Case "HP-100", "HP100"
                                        options(11) = "HP100"
                                    Case "HP-200", "HP-200LX", "HP200LX", "HP200"
                                        options(11) = "HP200"
                                    Case "HUSKY"
                                        options(11) = "HUSKY"
                                    Case "PC", "PC80", "2000"
                                    Case Else
                                        errorcode = 11
                                        errormessage = "Expect HP, HP100, HP200, PC or PC80 computer type."
                                        Exit Do
                                End Select

                            Case "SQID"
                                '-------------------------------------
                                'Is this a Laq/CC/Cagny type setup
                                '-------------------------------------
                                options(12) = Trim(dt)

                            Case "PRISMDEF"
                                If npoles = UBound(poledata, 1) Then
                                    errorcode = 93
                                    errormessage = "The number of poles is limited to" + Str(UBound(poledata, 2))
                                    Exit Do
                                End If
                                ts = dt
                                d = InStr(ts, " ")
                                If d = 0 Then
                                    errorcode = 90
                                    errormessage = "Missing prism name or height."
                                    Exit Do
                                End If
                                npoles = npoles + 1
                                polename(npoles) = UCase(Trim(Left(ts, d - 1)))
                                ts = LTrim(Mid(ts, d))
                                d = InStr(ts, " ")
                                If d = 0 Then
                                    errorcode = 91
                                    errormessage = "Missing prism offset."
                                    npoles = npoles - 1
                                    Exit Do
                                End If
                                poledata(npoles, 1) = CSng(Left(ts, d - 1))
                                poledata(npoles, 2) = CSng(Mid(ts, d + 1))

                            Case "DATUMDEF"
                                If ndatums = UBound(datum, 1) Then
                                    errorcode = 93
                                    errormessage = "The number of datums is limited to" + Str(UBound(datum, 1))
                                    Exit Do
                                End If
                                ts = dt
                                f = 0
                                Do Until Len(ts) = 0
                                    e = InStr(ts, " ")
                                    If e = 0 Then e = Len(ts) + 1
                                    f = f + 1
                                    temp(f) = Left(ts, e - 1)
                                    ts = LTrim(Mid(ts, e))
                                Loop
                                If f <> 5 Then
                                    errorcode = 94
                                    errormessage = "Format DATUMDEF as Name X Y Z Height"
                                    Exit Do
                                Else
                                    ndatums = ndatums + 1
                                    datum(ndatums, 1) = CDbl(temp(2))
                                    datum(ndatums, 2) = CDbl(temp(3))
                                    datum(ndatums, 3) = CDbl(temp(4))
                                    datum(ndatums, 4) = CDbl(temp(5))
                                    datumname(ndatums) = temp(1)

                                End If

                                '------------------------------------------
                                'The unit defaults file contains the default
                                'information for each unit.
                                '------------------------------------------
                            Case "UNITFILE"

                                '------------------------------------------
                                'These are the fields stored in the unit
                                'defaults file.  Make a packed string of them
                                '------------------------------------------
                            Case "UNITFIELDS"
                                options(17) = ""
                                dt = UCase(dt)
                                dt = Replace(dt, ",SUFFIX", "")
                                Do Until Len(dt) = 0
                                    e = InStr(dt, " ")
                                    If e = 0 Then e = Len(dt) + 1
                                    If options(17) = "" Then
                                        options(17) = Left(dt, e - 1)
                                    Else
                                        options(17) = options(17) + "," + Left(dt, e - 1)
                                    End If
                                    dt = LTrim(Mid(dt, e))
                                Loop

                                If options(17) <> "" Then
                                    b = InStr(options(17), ",")
                                    If b = 0 Then
                                        unitname = Trim(options(17))
                                    Else
                                        unitname = Trim(leftstr(options(17), b - 1))
                                    End If
                                Else
                                    unitname = "unit"
                                End If

                                '------------------------------------------
                                'Set limits for units
                                '------------------------------------------
                            Case "LIMIT"
                                dt = UCase(dt)
                                f = 0
                                Do Until Len(dt) = 0
                                    e = InStr(dt, " ")
                                    If e = 0 Then e = Len(dt) + 1
                                    f = f + 1
                                    temp(f) = Left(dt, e - 1)
                                    dt = LTrim(Mid(dt, e))
                                Loop
                                tempunitname = temp(1)
                                For a = 1 To unitlimits
                                    If Trim(limitname(a)) = tempunitname Then
                                        errorcode = 95
                                        errormessage = tempunitname + " defined twice."
                                        Exit Do
                                    End If
                                Next a
                                Select Case temp(2)
                                    Case "RECT"
                                        If f <> 6 Then
                                            errorcode = 94
                                            errormessage = "Format RECT as Name RECT X1 Y1 X2 Y2"
                                            Exit Do
                                        End If
                                        unitlimits = unitlimits + 1
                                        limitname(unitlimits) = tempunitname
                                        limits(unitlimits, 1) = CDbl(temp(3))
                                        limits(unitlimits, 2) = CDbl(temp(4))
                                        limits(unitlimits, 3) = CDbl(temp(5))
                                        limits(unitlimits, 4) = CDbl(temp(6))

                                    Case Else
                                        errormessage = "Missing unit type."
                                        errorcode = 81
                                        Exit Do

                                End Select

                                '------------------------------------------
                                ' Optionally put the sitename here.
                                '------------------------------------------
                            Case "SITE"
                                options(18) = dt

                                '------------------------------------------
                                ' For now this variable just indicates
                                ' whether there is a printer.  In the
                                ' future, it may say what kind of printer
                                ' (ie. ir printer or normal printer.
                                '------------------------------------------
                            Case "PRINTER"
                                options(19) = UCase(dt)

                            Case "TAMISRATE"

                            Case "VHDORXYZ"
                                options(21) = UCase(dt)

                                '--------------------------------------------------------
                                ' If we reach this point, the line must begin with either
                                ' a variable or unit name.  Otherwise it is an error.
                                '--------------------------------------------------------
                            Case Else

                                errorflag = 1
                                '-------------------------------------
                                ' First check to see if its an already
                                ' defined variable.
                                '-------------------------------------
                                For c = 1 To vars

                                    'check if defined variable

                                    If varlist(c) = key Then
                                        b = InStr(dt, " ") ' look for space
                                        If b = 0 Then b = Len(dt) + 1 'no value list

                                        '-------------------------------------
                                        'variable name and data
                                        '-------------------------------------
                                        varn = UCase(LTrim(RTrim(Left(dt, b - 1))))
                                        dat = LTrim(RTrim(Mid(dt, b + 1)))

                                        '-------------------------------------
                                        ' check for possible commands
                                        '-------------------------------------
                                        Select Case varn
                                            Case "INPUT"
                                                Select Case UCase(dat)
                                                    Case "TEXT"
                                                        vtype(c) = 1
                                                    Case "NUMERIC"
                                                        vtype(c) = 2
                                                    Case "INSTRUMENT"
                                                        vtype(c) = 3
                                                    Case "MENU"
                                                        vtype(c) = 4
                                                    Case "UNIT"
                                                        vtype(c) = 5

                                                        '---------------------
                                                        'For CC system only.
                                                        '---------------------
                                                    Case "SQID"
                                                        vtype(c) = 6
                                                        vlen(c) = 11
                                                        If options(10) = "" Then
                                                            errorcode = 7
                                                            errormessage = "SQIDFILE option must be specified first."
                                                            Exit Do
                                                        End If

                                                    Case Else
                                                        errorcode = 7
                                                        errormessage = Left(dat, 20) + " is an unrecognized INPUT type. "
                                                        Exit Do
                                                End Select

                                            Case "DEFAULT"
                                                vardefault(c) = dat

                                            Case "PRINT"
                                                dat = UCase(Left(dat, 1))
                                                If dat = "Y" Then varprint(c) = True

                                            Case "PROMPT"
                                                varprompt(c) = dat
                                                If Len(varprompt(c)) > 60 Then varprompt(c) = Left(varprompt(c), 60)

                                            Case "MENULIST"
                                                vmenu(c) = ""
                                                Do Until Len(dat) = 0
                                                    e = InStr(dat, " ")
                                                    If e = 0 Then e = Len(dat) + 1
                                                    If vmenu(c) = "" Then
                                                        vmenu(c) = UCase(Left(dat, e - 1))
                                                    Else
                                                        vmenu(c) = vmenu(c) + "," + UCase(Left(dat, e - 1))
                                                    End If
                                                    dat = LTrim(Mid(dat, e))
                                                Loop

                                            Case "VARLEN"
                                                vlen(c) = CInt(dat)

                                            Case "CARRY"
                                                vcarry(c) = True

                                            Case "INCREMENT"
                                                vincr(c) = True

                                            Case "REQUIRED"
                                                varrequired(c) = True

                                            Case "VARLOC"

                                            Case Else
                                                errorcode = 3
                                                errormessage = "Unrecognized command " + varn + " "
                                                Exit Do
                                        End Select

                                        errorflag = 0
                                        Exit For
                                    End If
                                Next c
                                If errorflag = 1 Then
                                    errorcode = 2
                                    errormessage = "No FIELD statement or unrecognized command " + key + ". "
                                    Exit Do
                                End If
                        End Select
                    Else
                        errorcode = 5
                        errormessage = "Unrecognized statement or field."
                        Exit Do
                    End If
                End If
            Loop
            file2.Close()
            file2.Dispose()

            If errorcode = 0 Then
                Cursor.Current = Cursors.Default
                answer = MsgBox("This file will be converted to EDM-Mobile format.", MsgBoxStyle.OkCancel, "EDM-Mobile")
                If answer = 1 Then
                    needtoconvert = 1
                    Cursor.Current = Cursors.WaitCursor
                Else
                    errorcode = 100
                    errormessage = "The configuration file must be converted to EDM-Mobile format for this program to work."
                End If
            End If

        Else

            iniclass = "[EDM]"
            inidata(1, 1) = "Sitename"
            inidata(2, 1) = "Database"
            inidata(3, 1) = "PointTable"
            'inidata(4, 1) = "Instrument"
            'inidata(5, 1) = "COMport"
            'inidata(6, 1) = "COMparameters"
            inidata(7, 1) = "SQID"
            inidata(8, 1) = "Unitfields"
            inidata(9, 1) = "Limitchecking"
            'inidata(10, 1) = "GridFields"
            inidata(11, 1) = "WriteLogFile"

            Call ReadIni(cfgfile, iniclass, inidata, status)

            If inidata(1, 2) <> "" Then
                options(18) = inidata(1, 2)
            Else
                options(18) = "EDM"
            End If

            If inidata(2, 2) <> "" Then
                options(1) = inidata(2, 2)
            Else
                options(1) = "EDM"
            End If

            If inidata(3, 2) <> "" Then
                options(22) = inidata(3, 2)
            Else
                options(22) = "EDM"
            End If

            'If inidata(4, 2) <> "" Then
            '    Select Case UCase(inidata(4, 2))
            '    Case "TOPCON", "SOKKIA", "WILD", "NONE"
            '        options(4) = inidata(4, 2)
            '    Case Else
            '        errorcode = 21
            '        errormessage = "Unrecognized EDM type.  Recognized types are Topcon, Wild, Sokkia and None."
            '    End Select
            'Else
            '    options(4) = "NONE"
            'End If

            'inidata(5, 1) = "COMport"
            'options(3) = inidata(5, 2)
            'inidata(6, 1) = "COMparameters"
            'options(2) = inidata(6, 2)

            If LCase(inidata(7, 2)) = "yes" Then options(12) = "Yes"

            If LCase(inidata(9, 2)) = "yes" Then use_limitchecking = True Else use_limitchecking = False
            If LCase(inidata(11, 2)) = "yes" Then WriteLogFile = True Else WriteLogFile = False

            If inidata(8, 2) <> "" Then
                options(17) = inidata(8, 2)
            ElseIf inidata(10, 2) <> "" Then
                options(17) = inidata(10, 2)
            Else
                options(17) = ""
            End If
            options(17) = UCase(options(17))
            options(17) = Replace(options(17), ",SUFFIX", "")

            ' Do some checking on unit fields
            ' Remove all fields that are default fields anyway
            Dim t As String
            Dim t2 As String
            Dim f1 As String
            t = options(17)
            t2 = ""
            Do Until t = ""
                a = InStr(t, ",")
                If a = 0 Then
                    f1 = t
                    t = ""
                Else
                    f1 = Left(t, a - 1)
                    t = Mid(t, a + 1)
                End If
                Select Case LCase(f1)
                    Case "x", "y", "z", "date", "time", "hangle", "vangle"
                    Case Else
                        If t2 = "" Then
                            t2 = f1
                        Else
                            t2 = t2 + "," + f1
                        End If
                End Select
            Loop
            options(17) = t2

            unitname = "unit"

            'now need to load variables
            'First, need to read through CFG looking for the field names

            Dim file3 As New System.IO.StreamReader(cfgfile)
            vars = 0
            Do
                ts = file3.ReadLine
                If ts Is Nothing Then Exit Do
                ts = Trim(ts)
                If leftstr(ts, 1) = "[" Then
                    ts = UCase(ts)
                    Select Case ts
                        Case "[EDM]", "[BUTTON1]", "[BUTTON2]", "[BUTTON3]", "[BUTTON4]", "[BUTTON5]", "[BUTTON6]"
                        Case Else
                            vars = vars + 1
                            varlist(vars) = ts
                    End Select
                End If
            Loop
            file3.Close()
            file3.Dispose()

            For a = 1 To 100
                inidata(a, 1) = ""
                inidata(a, 2) = ""
            Next a
            inidata(1, 1) = "Type"
            inidata(2, 1) = "Prompt"
            inidata(3, 1) = "Menu"
            inidata(4, 1) = "Length"
            inidata(5, 1) = "Increment"
            inidata(6, 1) = "Carry"
            inidata(7, 1) = "Unique"
            inidata(8, 1) = "Required"

            For c = 1 To vars

                For a = 1 To 8
                    inidata(a, 2) = ""
                Next a
                iniclass = varlist(c)
                Call ReadIni(cfgfile, iniclass, inidata, status)

                'strip the [ ]
                varlist(c) = Mid(varlist(c), 2)
                varlist(c) = leftstr(varlist(c), Len(varlist(c)) - 1)

                Select Case LCase(inidata(1, 2))
                    Case "text"
                        vtype(c) = 1
                    Case "numeric"
                        vtype(c) = 2
                    Case "instrument"
                        vtype(c) = 3
                    Case "menu"
                        vtype(c) = 4
                    Case "unit"
                        vtype(c) = 5
                    Case Else
                End Select

                varprompt(c) = inidata(2, 2)
                vmenu(c) = inidata(3, 2)

                If inidata(4, 2) <> "" Then vlen(c) = CInt(inidata(4, 2))

                If LCase(inidata(5, 2)) = "true" Or LCase(inidata(5, 2)) = "yes" Then vincr(c) = True Else vincr(c) = False

                If LCase(inidata(6, 2)) = "true" Or LCase(inidata(6, 2)) = "yes" Then vcarry(c) = True Else vcarry(c) = False

                If LCase(inidata(7, 2)) = "true" Or LCase(inidata(7, 2)) = "yes" Then vunique(c) = True Else vunique(c) = False

                If LCase(inidata(8, 2)) = "true" Or LCase(inidata(8, 2)) = "yes" Then varrequired(c) = True Else varrequired(c) = False
            Next c

        End If

    End Sub

    Sub write_datafile_ini(ByVal filename As String, ByRef status As Integer)

        'This routine writes an ini for the data table within a site

        Dim inidata(100, 2) As String
        Dim iniclass As String
        Dim a As Integer
        Dim b As Integer
        Dim wstatus As Integer

        If LCase(Right(filename, 4)) <> ".cfg" Then filename = filename + ".cfg"

        'If it exists, kill it and start over
        If File.Exists(filename) Then File.Delete(filename)

        'Write general ini variables
        iniclass = "[EDM]"
        inidata(1, 1) = "Sitename"
        inidata(1, 2) = options(18)
        inidata(2, 1) = "Database"
        inidata(2, 2) = options(1)
        inidata(3, 1) = "PointTable"
        inidata(3, 2) = options(22)
        inidata(4, 1) = "SQID"
        inidata(4, 2) = options(12)
        inidata(5, 1) = "Unitfields"
        inidata(5, 2) = options(17)
        inidata(6, 1) = "Limitchecking"
        If use_limitchecking = True Then inidata(6, 2) = "Yes" Else inidata(6, 2) = "No"

        inidata(7, 1) = "Total Station"
        inidata(7, 2) = options(4)
        inidata(8, 1) = "COM"
        inidata(8, 2) = options(2)
        inidata(9, 1) = "Settings"
        inidata(9, 2) = options(3)
        inidata(10, 1) = "CurrentStation"
        inidata(10, 2) = currentstationname
        inidata(11, 1) = "CurrentStationX"
        inidata(11, 2) = CStr(currentstationx)
        inidata(12, 1) = "CurrentStationY"
        inidata(12, 2) = CStr(currentstationy)
        inidata(13, 1) = "CurrentStationZ"
        inidata(13, 2) = CStr(currentstationz)
        inidata(14, 1) = "Referencedatum"
        inidata(14, 2) = referencedatum


        'Write general info for this config
        Call WriteIni(filename, iniclass, inidata, wstatus)

        'Clear the arrays
        For b = 1 To UBound(inidata, 1)
            inidata(b, 1) = ""
            inidata(b, 2) = ""
        Next b

        'Write out the buttons
        Dim frmmain As New frmMain
        If frmmain.Button1.Visible = True Then Call save_button(1, frmmain.Button1.Text, filename)
        If frmmain.Button2.Visible = True Then Call save_button(2, frmmain.Button2.Text, filename)
        If frmmain.Button3.Visible = True Then Call save_button(3, frmmain.Button3.Text, filename)
        If frmmain.Button4.Visible = True Then Call save_button(4, frmmain.Button4.Text, filename)
        If frmmain.Button5.Visible = True Then Call save_button(5, frmmain.Button5.Text, filename)
        If frmmain.Button6.Visible = True Then Call save_button(6, frmmain.Button6.Text, filename)

        'Now write section for each field
        For a = 1 To vars
            Call savevar(a, filename)
        Next a

    End Sub
    Sub WriteIni(ByVal inifile As String, ByVal iniclass As String, ByVal inidata(,) As String, ByRef status As Integer)

        Dim Gotit As Boolean
        Dim a1 As String
        Dim a2 As String
        Dim iclass As Integer
        Dim c As Integer
        Dim b As Integer
        Dim varname As String
        Dim vardata As String
        Dim flag As Integer
        Dim tempfile As String = "temp" + hash(4) + ".tmp"

        If inifile = "" Or inifile Is Nothing Then Exit Sub

        If Not File.Exists(inifile) Then
            Dim file3 As StreamWriter = File.CreateText(inifile)
            file3.WriteLine(iniclass)
            file3.Flush()
            file3.Close()
            file3.Dispose()
        End If

        Do
            Dim file1 As New StreamReader(inifile)
            If File.Exists(tempfile) Then
                Dim t2 As Double = Timer
                Try
                    File.Delete(tempfile)
                Catch ex As Exception
                    If Timer - t2 > 2 Then Exit Try
                End Try
            End If
            Dim file2 As StreamWriter = File.CreateText(tempfile)

            Do

                a1 = file1.ReadLine
                If a1 Is Nothing Then Exit Do
                a2 = UCase(Trim(a1))

                If Left(a2, 1) = "[" Then
                    file2.WriteLine(a1)
                    If a2 = UCase(iniclass) Then
                        iclass = 1
                        Gotit = True
                        For c = 1 To UBound(inidata, 1)
                            If Trim(inidata(c, 1)) <> "" And Trim(inidata(c, 2)) <> "" Then
                                file2.WriteLine(inidata(c, 1) + "=" + CStr(inidata(c, 2)))
                            End If
                        Next c
                    Else
                        iclass = 0
                    End If

                ElseIf iclass = 1 Then
                    b = InStr(a2, "=")
                    If b <> 0 Then
                        varname = UCase(Left(a2, b - 1))
                        vardata = Mid(a2, b + 1)
                        flag = 0
                        For c = 1 To UBound(inidata, 1)
                            If UCase(inidata(c, 1)) = varname Then
                                flag = 1
                                Exit For
                            End If
                        Next c
                        If flag = 0 Then
                            file2.WriteLine(a1)
                        End If
                    Else
                        file2.WriteLine(a1)
                    End If
                Else
                    file2.WriteLine(a1)
                End If
            Loop

            file1.Close()
            file1.Dispose()

            file2.Flush()
            file2.Close()
            file2.Dispose()

            'If the section was never found, add it and start over
            If Not Gotit Then
                Dim file4 As StreamWriter = File.AppendText(inifile)
                file4.WriteLine(" ")
                file4.WriteLine(iniclass)
                file4.Flush()
                file4.Close()
                file4.Dispose()
            End If

        Loop Until Gotit = True

        'Delete the existing one, rename the new one as the existing one
        Dim t1 As Double = Timer
        Try
            File.Delete(inifile)
        Catch ex As Exception
            If Timer - t1 > 2 Then Exit Try
        End Try
        File.Move(tempfile, inifile)
        File.Delete(tempfile)

    End Sub

    Sub ReadIni(ByVal inifile As String, ByVal iniclass As String, ByVal inidata(,) As String, ByRef status As Integer)

        Dim iclass As Integer
        Dim b As Integer
        Dim c As Integer
        Dim varname As String
        Dim vardata As String
        Dim a1 As String
        Dim a2 As String

        'If it does not exist, create it and add the iniclass empty heading
        If Not File.Exists(inifile) Then
            Dim file2 As StreamWriter = File.CreateText(inifile)
            file2.WriteLine(iniclass)
            file2.Close()
        End If

        'Now read through the file and fill in the values from the ini
        Dim file1 As New StreamReader(inifile)
        Do
            a1 = file1.ReadLine
            If a1 Is Nothing Then Exit Do
            a1 = Trim(a1)
            a2 = UCase(a1)
            If Left(a2, 1) = "[" Then
                If a2 = UCase(iniclass) Then
                    iclass = 1
                Else
                    iclass = 0
                End If
            ElseIf iclass = 1 Then
                b = InStr(a2, "=")
                If b <> 0 Then
                    varname = Left(a2, b - 1)
                    vardata = Mid(a1, b + 1)
                    For c = 1 To UBound(inidata, 1)
                        If UCase(inidata(c, 1)) = UCase(varname) Then
                            If inidata(c, 2) = "" Then
                                inidata(c, 2) = vardata
                            Else
                                inidata(c, 2) = inidata(c, 2) + Chr(1) + vardata
                            End If
                            Exit For
                        End If
                    Next c
                End If
            End If
        Loop

        file1.Close()
        file1.Dispose()

    End Sub

    Sub create_db(ByVal dbname As String)

        If File.Exists(dbname) Then
            File.Delete(dbname)
        End If

        If Not Directory.Exists(Path.GetDirectoryName(Path.GetFullPath(dbname))) Then
            dbname = Path.GetDirectoryName(cfgfile) + "\" + Path.GetFileName(dbname)
        End If

        Dim connStr As String = "Data Source = " + dbname

        Dim engine As New SqlCeEngine(connStr)
        engine.CreateDatabase()
        engine.Dispose()

    End Sub

    Sub add_field_tb()

        Dim a As Integer
        Dim sql(100) As String
        Dim sqlcmd As String

        sql(0) = "CREATE TABLE " & options(22) & " (RECNO int primary key,"
        For a = 1 To vars
            If UCase(varlist(a)) = "DATE" Or UCase(varlist(a)) = "DAY" Then
                sql(a) = varlist(a) & " datetime"     'adDouble
            ElseIf UCase(varlist(a)) = "TIME" Then
                sql(a) = "TIME datetime"       'adDouble
            Else
                Select Case vtype(a)
                    Case 1
                        sql(a) = varlist(a) & " nvarchar(" & CStr(vlen(a)) & ")" 'adVarWChar
                    Case 2
                        sql(a) = varlist(a) & " float"     'adDouble
                    Case 3
                        sql(a) = varlist(a) & " float"     'adDouble
                    Case 4
                        sql(a) = varlist(a) & " nvarchar(" & CStr(vlen(a)) & ")" 'adVarWChar
                    Case 5
                        sql(a) = varlist(a) & " nvarchar(" & CStr(vlen(a)) & ")" 'adVarWChar
                    Case Else
                End Select
            End If
        Next a

        sqlcmd = sql(0)
        For a = 1 To vars - 1
            sqlcmd = sqlcmd & sql(a) & ","
        Next
        sqlcmd = sqlcmd & sql(vars) & ")"

        Dim connStr As String = "Data Source = '" & options(1) & "'"
        Dim conn As SqlCeConnection = Nothing

        Try
            conn = New SqlCeConnection(connStr)
            conn.Open()

            Dim cmd As SqlCeCommand = conn.CreateCommand()
            cmd.CommandText = sqlcmd
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE INDEX Recno ON " & options(22) & " (Recno)"
            cmd.ExecuteNonQuery()

        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error creating field data table (" & options(22) & ") in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Sub

        Finally
            conn.Close()

        End Try

    End Sub

    
    Sub parse_list(ByVal ltext As String, ByVal d As String, ByRef n As Integer, ByVal l() As String)

        Dim t1 As String
        Dim a As Integer

        t1 = Trim(ltext)

        For a = 1 To UBound(l, 1)
            l(a) = ""
        Next a

        n = 0
        Do Until t1 = "" Or n = UBound(l, 1)
            a = InStr(t1, d)
            If a <> 0 Then
                If Left(t1, a - 1) <> "" Then
                    n = n + 1
                    l(n) = Left(t1, a - 1)
                End If
                t1 = Mid(t1, a + 1)
            ElseIf t1 <> "" Then
                n = n + 1
                l(n) = t1
                Exit Do
            End If
        Loop

    End Sub

    Sub fix_cfg_info()

        Dim a As Integer
        Dim flag As Boolean

        If options(1) = "" Then
            options(1) = fpath + "edm.sdf"
        End If

        If options(22) = "" Then
            options(22) = "edm"
        End If

        If InStr(options(1), ".") <> 0 Then
            a = InStr(options(1), ".")
            If LCase(Mid(options(1), a + 1)) <> "sdf" Then
                options(1) = Left(options(1), a) + "sdf"
            End If
        End If

        If InStr(options(1), "\") = 0 Then
            options(1) = fpath + options(1)
        End If

        For a = 1 To vars
            Select Case varlist(a)
                Case "X", "Y", "Z", "VANGLE", "HANGLE", "SLOPED", "DATUMX", "DATUMY", "DATUMZ"
                    vtype(a) = 2
                    vlen(a) = 13
                Case "SUFFIX"
                    vtype(a) = 2
                    vlen(a) = 4
                Case "PRISM"
                    vtype(a) = 2
                    vlen(a) = 8
                Case "DATUMNAME"
                    vtype(a) = 1
                    vlen(a) = 20
                Case "DAY", "TIME", "DATE"
                    vtype(a) = 1
                    vlen(a) = 14
                Case Else
                    If vtype(a) = 0 Then vtype(a) = 1
            End Select
        Next a

        flag = False
        For a = 1 To vars
            If UCase(varlist(a)) = "UNIT" Then
                flag = True
                Exit For
            End If
        Next a

        If Not flag Then
            vars = vars + 1
            varlist(vars) = "UNIT"
            vtype(vars) = 1
            vlen(vars) = 6
            Call savevar(vars, cfgfile)
        End If

        flag = False
        For a = 1 To vars
            If UCase(varlist(a)) = "ID" Then
                flag = True
                Exit For
            End If
        Next a

        If Not flag Then
            vars = vars + 1
            varlist(vars) = "ID"
            vtype(vars) = 1
            vlen(vars) = 5
            Call savevar(vars, cfgfile)
        End If

        flag = False
        For a = 1 To vars
            If UCase(varlist(a)) = "SUFFIX" Then
                flag = True
                Exit For
            End If
        Next a

        If Not flag Then
            vars = vars + 1
            varlist(vars) = "SUFFIX"
            vtype(vars) = 2
            vlen(vars) = 2
            Call savevar(vars, cfgfile)
        End If

        flag = False
        For a = 1 To vars
            If UCase(varlist(a)) = "PRISM" Then
                flag = True
                Exit For
            End If
        Next a

        If Not flag Then
            vars = vars + 1
            varlist(vars) = "PRISM"
            vtype(vars) = 2
            vlen(vars) = 8
            Call savevar(vars, cfgfile)
        End If

        flag = False
        For a = 1 To vars
            If UCase(varlist(a)) = "X" Then
                flag = True
                Exit For
            End If
        Next a

        If Not flag Then
            vars = vars + 1
            varlist(vars) = "X"
            vtype(vars) = 2
            vlen(vars) = 13
            Call savevar(vars, cfgfile)
        End If

        flag = False
        For a = 1 To vars
            If UCase(varlist(a)) = "Y" Then
                flag = True
                Exit For
            End If
        Next a

        If Not flag Then
            vars = vars + 1
            varlist(vars) = "Y"
            vtype(vars) = 2
            vlen(vars) = 13
            Call savevar(vars, cfgfile)
        End If

        flag = False
        For a = 1 To vars
            If UCase(varlist(a)) = "Z" Then
                flag = True
                Exit For
            End If
        Next a

        If Not flag Then
            vars = vars + 1
            varlist(vars) = "Z"
            vtype(vars) = 2
            vlen(vars) = 13
            Call savevar(vars, cfgfile)
        End If

        For a = 1 To vars
            If (vtype(a) = 1 Or vtype(a) = 4) And vlen(a) = 0 Then
                vlen(a) = 10
            End If
        Next a

        For a = 1 To vars
            Select Case varlist(a)
                Case "VANGLE"
                    If varprompt(a) = "" Then varprompt(a) = "Vert. angle"
                Case "HANGLE"
                    If varprompt(a) = "" Then varprompt(a) = "Hor. angle"
                Case "SLOPED"
                    If varprompt(a) = "" Then varprompt(a) = "Slope dist."
                Case Else
                    If varprompt(a) = "" Then varprompt(a) = UCase(Left(varlist(a), 1)) + LCase(Mid(varlist(a), 2))
            End Select
        Next a

        'This next bit of code makes sure that the unit fields include UNIT and ID
        'and are written in such a way that the program knows what to do.
        If options(17) <> "" Then
            Dim prevunitfields As String = options(17)
            Dim unitfields(100) As String
            Dim unitcount As Integer
            Call parse_list(options(17), ",", unitcount, unitfields)
            options(17) = "UNIT,ID"
            For a = 1 To unitcount
                Select Case LCase(unitfields(a))
                    Case "id", "unit"
                    Case Else
                        options(17) = options(17) + "," + unitfields(a)
                End Select
            Next
            If options(17) <> prevunitfields Then
                Dim inidata(1, 2) As String
                Dim iniclass As String = "[EDM]"
                Dim wstatus As Integer = 0
                inidata(1, 1) = "Unitfields"
                inidata(1, 2) = options(17)
                Call WriteIni(cfgfile, iniclass, inidata, wstatus)
            End If
        End If

    End Sub

    Sub table_status(ByVal tbname As String, ByRef status As Integer)

        'table does not exist
        status = -1

        If options(1) = "" Then Exit Sub
        If tbname = "" Then Exit Sub

        'first see if the table exists
        Dim systemtables As New SqlCeConnection("Datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = systemtables.CreateCommand
        sql.CommandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='" + tbname + "' AND TABLE_TYPE = 'TABLE'"
        Dim r As Object
        Dim rdr As SqlCeDataReader
        '
        Try
            systemtables.Open()
        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Could not open " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Sub
        Finally
        End Try

        Try
            r = sql.ExecuteScalar
            If LCase(CStr(r)) = LCase(tbname) Then status = 0 'table exists but is empty
        Catch
            MsgBox("Reading table schema failed." & vbNewLine & Err.Description, MsgBoxStyle.Exclamation, "Issue")
            Exit Sub
        Finally
        End Try

        'If it does, see how many records it has
        If status <> -1 Then
            sql.CommandText = "SELECT * FROM " + tbname
            Try
                rdr = sql.ExecuteReader
                While rdr.Read()
                    status = status + 1 'table exists and has records
                End While
                rdr.Close()
            Catch sqlex As System.Data.SqlServerCe.SqlCeException
                MsgBox("Table record count failed." & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
                Exit Try
            Finally
            End Try
        End If
        systemtables.Close()

    End Sub

    Sub add_datum_table()

        Dim connStr As String = "Data Source = " + options(1)
        Dim conn As SqlCeConnection = Nothing

        Try
            conn = New SqlCeConnection(connStr)
            conn.Open()

            Dim cmd As SqlCeCommand = conn.CreateCommand()
            cmd.CommandText = "CREATE TABLE edm_datums (name nvarchar(20), x float, y float, z float, day datetime, time datetime)"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE INDEX datumname ON edm_datums (name)"
            cmd.ExecuteNonQuery()

        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error creating datum table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Sub

        Finally
            conn.Close()

        End Try

    End Sub

    Sub add_prism_table()

        Dim connStr As String = "Data Source = " + options(1)
        Dim conn As SqlCeConnection = Nothing

        Try
            conn = New SqlCeConnection(connStr)
            conn.Open()

            Dim cmd As SqlCeCommand = conn.CreateCommand()
            cmd.CommandText = "CREATE TABLE edm_poles (name nvarchar(20), height float, offset float)"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE INDEX polename ON edm_poles (name)"
            cmd.ExecuteNonQuery()

        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error creating prism table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Sub

        Finally
            conn.Close()

        End Try

    End Sub

    Sub add_unit_table()

        Dim sqlcmd As String
        Dim a As Integer
        Dim sql1 As String = ""
        Dim fieldlen_unit As Integer
        Dim fieldlen_id As Integer

        Dim connStr As String = "Data Source = " + options(1)
        Dim conn As SqlCeConnection = Nothing

        Call getfieldlen("unit", fieldlen_unit)
        If fieldlen_unit = 0 Then fieldlen_unit = 10
        Call getfieldlen("id", fieldlen_id)
        If fieldlen_id = 0 Then fieldlen_id = 10

        sqlcmd = "CREATE TABLE edm_units (unit nvarchar(" & CStr(fieldlen_unit) & "), id nvarchar(" & CStr(fieldlen_id) & "), minx float, miny float, maxx float, maxy float, centerx float, centery float, radius float "

        'Go through the CFG and add fields that could be used as unit fields
        For a = 1 To vars
            Select Case LCase(varlist(a))
                Case "x", "y", "z", "datumname", "datumx", "datumy", "datumz", "day", "date", "time", "unit", "suffix", "id", "hangle", "vangle", "prism"
                Case Else
                    Select Case vtype(a)
                        Case 1
                            sql1 = sql1 + "," + varlist(a) & " nvarchar(" & CStr(vlen(a)) & ")" 'adVarWChar
                        Case 2
                            sql1 = sql1 + "," + varlist(a) & " float"     'adDouble
                        Case 3
                            sql1 = sql1 + "," + varlist(a) & " float"     'adDouble
                        Case 4
                            sql1 = sql1 + "," + varlist(a) & " nvarchar(" & CStr(vlen(a)) & ")" 'adVarWChar
                        Case 5
                            sql1 = sql1 + "," + varlist(a) & " nvarchar(" & CStr(vlen(a)) & ")" 'adVarWChar
                        Case Else
                    End Select
            End Select
        Next a

        sqlcmd = sqlcmd + sql1 + ")"
        
        Try
            conn = New SqlCeConnection(connStr)
            conn.Open()

            Dim cmd As SqlCeCommand = conn.CreateCommand()
            cmd.CommandText = sqlcmd
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE INDEX Unitname ON edm_units (Unit)"
            cmd.ExecuteNonQuery()

        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error creating unit table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Sub

        Finally
            conn.Close()

        End Try

    End Sub

    Function leftstr(ByVal t As String, ByVal l As Integer) As String

        'This kludge is to work around the fact that CE doesn't accept Left on forms

        If l > 0 Then leftstr = Left(t, l) Else leftstr = ""

    End Function

    Sub use_defaults()

        Dim status As Integer
        Dim a As Integer

        vars = 17

        For a = 1 To vars
            varlist(a) = ""
            vtype(a) = 0
            vlen(a) = 0
            vcarry(a) = False
            vincr(a) = False
            vmenu(a) = ""
            varprompt(a) = ""
        Next a

        varlist(1) = "Sitename"
        vtype(1) = 1
        vlen(1) = 10
        vcarry(1) = True

        varlist(2) = "Unit"
        vtype(2) = 4
        vmenu(2) = "UNIT1,UNIT2"
        vlen(2) = 5
        vcarry(2) = True

        varlist(3) = "ID"
        vtype(3) = 1
        vincr(3) = True
        vcarry(3) = True
        vlen(3) = 5

        varlist(4) = "Suffix"
        vtype(4) = 2
        vcarry(4) = True

        varlist(5) = "Code"
        vtype(5) = 4
        vmenu(5) = "STONE,BONE,TOOTH,TOPO"
        vlen(5) = 10
        vcarry(5) = True

        varlist(6) = "Level"
        vlen(6) = 10
        vmenu(6) = "Level1,Level2"
        vtype(6) = 4
        vcarry(6) = True

        varlist(7) = "Excavator"
        vlen(7) = 10
        vmenu(7) = "Utsav,Steve,Radu"
        vtype(7) = 4
        vcarry(7) = True

        varlist(8) = "X"
        varlist(9) = "Y"
        varlist(10) = "Z"
        varlist(11) = "PRISM"
        'varlist(10) = "HANGLE"
        'varlist(11) = "VANGLE"
        'varlist(12) = "SLOPED"
        varlist(12) = "DAY"
        varlist(13) = "TIME"
        varlist(14) = "DATUMNAME"
        varlist(15) = "DATUMX"
        varlist(16) = "DATUMY"
        varlist(17) = "DATUMZ"

        Call fix_cfg_info()

        use_limitchecking = False
        WriteLogFile = True

        Dim dbname As String
        dbname = options(1)
        unitname = "Unit"

        If InStr(dbname, "\") = 0 Then
            dbname = fpath + dbname
        End If

        If InStr(dbname, ".") = 0 Then
            dbname = dbname + ".sdf"
        End If

        options(1) = dbname
        options(17) = "UNIT,ID,LEVEL,EXCAVATOR"

        Call setup_dbs()

        If cfgfile = "" Then
            cfgfile = fpath + options(22) + ".cfg"
            LogFile = fpath + options(22) + "_log.txt"
        End If

        Call write_datafile_ini(cfgfile, status)

    End Sub

    Sub calculate_angle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByRef angle As Double)

        Dim rrun As Double
        Dim rise As Double
        Dim slope As Double
        Dim pi As Double = 3.14159265359
        Dim tangle As Double

        rrun = CDbl(x2) - CDbl(x1)
        rise = CDbl(y2) - CDbl(y1)

        If rrun = 0 Then
            tangle = 90
        Else
            slope = rise / rrun
            tangle = CSng(Math.Atan(slope) * 180.0# / pi)
        End If

        tangle = 90 - tangle

        If (rrun >= 0) And (rise >= 0) Then
            tangle = 0 + tangle
        ElseIf (rrun > 0) And (rise < 0) Then
            tangle = 360 + tangle
        ElseIf (rrun <= 0) And (rise >= 0) Then
            tangle = 180 + tangle
        ElseIf (rrun <= 0) And (rise < 0) Then
            tangle = 180 + tangle
        End If

        'tangle = tangle + 90
        If tangle >= 360 Then tangle = tangle - 360

        angle = tangle

    End Sub

    Sub conv_angle_to_degminsec(ByVal angle As Double, ByRef degrees As Integer, ByRef minutes As Integer, ByRef seconds As Integer)

        degrees = CInt(Int(angle))
        seconds = CInt(Int((angle - CDbl(degrees)) * 3600.0#))
        minutes = CInt(Int(seconds / 60))
        seconds = seconds Mod 60

    End Sub

    Sub record_point(ByRef status As Integer)

        Select Case options(4)
            Case "TOPCON", "WILD", "SOKKIA"

            Case "NONE"
                'Unwritten hand entry section was to go here

            Case "SIMULATION"

            Case Else
        End Select

    End Sub

    Function degtorad(ByVal degrees As Single) As Single

        degtorad = CSng(degrees * 0.0174532925199433)

    End Function

    Function radtodeg(ByVal rad As Double) As Double

        ' Note that this function is new (April 2025) which is why it uses double precision
        radtodeg = rad / (2.0 * System.Math.PI) * 360.0

    End Function

    Sub open_com()

        Dim errorcode As Integer

        If options$(4) <> "SIMULATION" Then

            If commcontrol.IsOpen Then commcontrol.Close()
            commcontrol.PortName = options(2)

            Dim comsettings(4) As String
            Dim n As Integer
            Call parse_list(options(3), ",", n, comsettings)
            commcontrol.BaudRate = CInt(comsettings(1))
            Select Case LCase(comsettings(2))
                Case "e"
                    commcontrol.Parity = Ports.Parity.Even
                Case "o"
                    commcontrol.Parity = Ports.Parity.Odd
                Case "n"
                    commcontrol.Parity = Ports.Parity.None
                Case Else
                    commcontrol.Parity = Ports.Parity.Space
            End Select
            commcontrol.DataBits = CInt(comsettings(3))
            Select Case comsettings(4)
                Case "0"
                    commcontrol.StopBits = Ports.StopBits.None
                Case "1"
                    commcontrol.StopBits = Ports.StopBits.One
                Case "2"
                    commcontrol.StopBits = Ports.StopBits.Two
            End Select

            commcontrol.ReadTimeout = 500
            commcontrol.WriteTimeout = 500
            commcontrol.ReceivedBytesThreshold = 1
            commcontrol.DtrEnable = True
            commcontrol.Handshake = Handshake.XOnXOff

            Try
                commcontrol.Open()

            Catch ex As Exception
                Cursor.Current = Cursors.Default
                MsgBox("Could not open " & options(2) & ":" & options(3) & " Make sure you are cabled and use the Station Comport Settings to change the settings.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End Try
        End If

        Call initedm()
        Call horizontal(errorcode)

    End Sub

    Sub close_com()

        If options$(4) <> "SIMULATION" Then
            If commcontrol.IsOpen Then commcontrol.Close()
        End If

    End Sub

    Sub edminput(ByVal a As String)

        Dim t As String = ""
        Dim timeone As Double = Microsoft.VisualBasic.DateAndTime.Timer

        a = ""
        If UCase(options$(4)) <> "SIMULATION" Then
            Do
                If commcontrol.IsOpen Then t = commcontrol.ReadExisting
                If t <> "" Then a = a + t
            Loop Until Right(a, 2) = Chr(13) + Chr(10) Or Microsoft.VisualBasic.DateAndTime.Timer - timeone > 7
        End If

    End Sub

    Sub edmoutput(ByVal d As String, ByRef errorcode As Integer)

        Dim term As String = Chr(13) + Chr(10)
        Dim bcc As String = ""
        Dim a As String
        Dim bTemp(200) As Byte
        'Dim Y As Integer

        errorcode = 0

        Select Case options(4)
            Case "TOPCON"
                Call makebcc(d, bcc)
                a = d + bcc + Chr(3) + term

            Case "WILD", "LEICA", "LEICA-GEO"
                a = d + term

            Case "SOKKIA"
                a = d

            Case Else
                Exit Sub

        End Select

        Select Case options(4)
            Case "TOPCON"
                If commcontrol.IsOpen Then commcontrol.Write(a)

            Case "WILD", "LEICA", "SOKKIA", "NIKON", "LEICA-GEO"
                If commcontrol.IsOpen Then commcontrol.Write(a)
                'frmmain.comm1.Output = a
                'Print #display.DataSource, a;

            Case Else
        End Select


    End Sub

    Sub horizontal(ByVal errorcode As Integer)

        Dim d As String = ""
        Dim a As String = ""

        errorcode = 0
        Select Case options(4)
            Case "TOPCON"
                d = "Z10"
                Call edmoutput(d, errorcode)
                Call edminput(a)
                If a = "CANCEL" Then errorcode = 27
            Case Else
        End Select

    End Sub

    Sub initedm()

        Dim errorcode As Integer = 0
        Dim d As String = ""
        Dim a As String = ""

        Select Case options(4)
            Case "TOPCON"
                d = "ST0"
                Call edmoutput(d, errorcode)

            Case "WILD", "LEICA"
                d = "SET/41/0"
                Call edmoutput(d, errorcode)
                If errorcode = 0 Then
                    Call edminput(a)
                End If
                d = "SET/149/2"
                Call edmoutput(d, errorcode)
                If errorcode = 0 Then
                    Call edminput(a)
                End If

            Case "LEICA-GEO"
                ' Do nothing

            Case Else
        End Select

    End Sub

    Sub makebcc(ByVal itext As String, ByRef otext As String)

        Dim b As Integer = 0
        Dim i As Integer
        Dim q As Integer
        Dim b1 As Integer
        Dim b2 As Integer

        For i = 1 To Len(itext)
            q = Asc(Mid(itext, i, 1))
            b1 = q And (Not b)
            b2 = b And (Not q)
            b = b1 Or b2
        Next i

        otext = LTrim(CStr(b))
        otext = Right("000" + otext, 3)

    End Sub

    Function rad_to_dms(ByVal angle As Double) As String

        Dim degrees As Integer
        Dim minutes As Integer
        Dim seconds As Integer

        Dim angle_dd As Double = radtodeg(angle)
        conv_angle_to_degminsec(angle_dd, degrees, minutes, seconds)
        rad_to_dms = CStr(degrees) + "." + Microsoft.VisualBasic.Right("00" + CStr(minutes), 2) + Microsoft.VisualBasic.Right("00" + CStr(seconds), 2)

    End Function

    Sub parsenez(ByVal nezdata As String, ByVal edmpoffset As Single, ByVal mesunits As String, ByVal angleunit As String, ByRef errorcode As Integer)

        Dim a As String
        Dim aa As Integer
        Dim bcc1 As String = ""
        Dim bcc2 As String = ""
        Dim d As String
        Dim i As Integer

        errorcode = 0
        Select Case options(4)
            Case "TOPCON"

                Do Until Asc(Left(nezdata, 1)) > 32 Or Len(nezdata) = 1
                    nezdata = Mid(nezdata, 2)
                Loop
                a = Left(nezdata, 1)
                If a <> "?" And a <> "R" Then
                    If a = "U" Then
                        errorcode = 5
                    Else
                        errorcode = 1
                    End If
                    Exit Sub
                End If

                aa = InStr(nezdata, Chr(3))
                If aa <> 0 Then
                    bcc1 = Mid(nezdata, aa - 3, 3)
                    d = Left(nezdata, aa - 4)
                    Call makebcc(d, bcc2)
                End If

                If bcc1 <> bcc2 Then
                    errorcode = 6
                Else
                    nezdata = LTrim(nezdata)
                    currentsloped = CDbl(FormatNumber(Mid(nezdata, 2, 9))) / 1000
                    currenthangle = FormatNumber(Mid(nezdata, 19, 4) + "." + Mid(nezdata, 23, 4), 4)
                    currentvangle = FormatNumber(Mid(nezdata, 12, 3) + "." + Mid(nezdata, 15, 4), 4)
                    edmpoffset = CSng(FormatNumber(Mid(nezdata, 43, 3))) / 1000
                    mesunits = Mid(nezdata, 11, 1)
                    angleunit = Mid(nezdata, 27, 1)
                    If angleunit <> "d" Then
                        errorcode = 2
                    End If
                End If

            Case "WILD", "LEICA"
                i = InStr(nezdata, "21.")
                If i = 0 Then
                    errorcode = 1
                    Exit Sub
                End If
                nezdata = Mid(nezdata, i)
                'IF LEN(nezdata) <> 64 OR a = 0 THEN
                '        errorcode = 1
                '        EXIT SUB
                'END IF

                If Mid(nezdata, 6, 1) <> "4" Then
                    errorcode = 2
                    Exit Sub
                End If

                ' This tests for hangle increasing left vs. right
                If Mid(nezdata, 7, 1) = "-" Then
                    errorcode = 200
                End If

                'Fixed to deal with Leica GSI 8 vs. 16
                nezdata = Mid(nezdata, 7)
                i = InStr(nezdata, " ")
                If i <> 0 Then
                    Dim hdata As String = Left(nezdata, i - 1)
                    currenthangle = FormatNumber(Left(hdata, Len(hdata) - 5) + "." + Right(hdata, 5), 4)
                Else
                    errorcode = 2
                    Exit Sub
                End If

                nezdata = Mid(nezdata, i + 1)
                If Mid(nezdata, 6, 1) <> "4" Then
                    errorcode = 2
                    Exit Sub
                End If

                nezdata = Mid(nezdata, 7)
                i = InStr(nezdata, " ")
                If i <> 0 Then
                    Dim vdata As String = Left(nezdata, i - 1)
                    currentvangle = FormatNumber(Left(vdata, Len(vdata) - 5) + "." + Right(vdata, 5), 4)
                Else
                    errorcode = 2
                    Exit Sub
                End If

                nezdata = Mid(nezdata, i + 1)
                If Left(nezdata, 5) <> "31..0" Then
                    errorcode = 2
                    Exit Sub
                End If

                i = InStr(nezdata, "+")
                If i <> 0 Then
                    nezdata = Mid(nezdata, i + 1)
                Else
                    errorcode = 2
                    Exit Sub
                End If

                i = InStr(nezdata, " ")
                If i <> 0 Then
                    Dim sdata As String = Left(nezdata, i - 1)
                    currentsloped = CDbl(FormatNumber(sdata)) / 1000
                Else
                    errorcode = 2
                    Exit Sub
                End If

                If currentsloped = 0 Then
                    errorcode = 2
                    Exit Sub
                End If

                nezdata = Mid(nezdata, i + 1)
                mesunits = Mid(nezdata, i + 5)

                i = InStr(nezdata, "+")
                If i <> 0 Then
                    nezdata = Mid(nezdata, i + 1)
                    i = InStr(nezdata, "+")
                    If i <> 0 Then
                        edmpoffset = CSng(FormatNumber(Mid(nezdata, i, 4))) / 1000
                    End If
                End If

            Case "LEICA-GEO"
                i = InStr(nezdata$, ":")
                If i = 0 Then
                    errorcode = 20
                    Exit Sub
                End If
                nezdata$ = Mid(nezdata$, i + 1)
                ' MsgBox(nezdata$, MsgBoxStyle.Information)
                i = InStr(nezdata, ",")
                If i = 0 Then
                    errorcode = 20
                    Exit Sub
                End If
                nezdata = Mid(nezdata, i + 1)
                i = InStr(nezdata, ",")
                If i = 0 Then
                    errorcode = 20
                    Exit Sub
                End If
                Dim hdata As String = Left(nezdata, i - 1)
                If hdata = "" Then
                    errorcode = 20
                    Exit Sub
                End If
                currenthangle = rad_to_dms(CDbl(hdata))
                nezdata = Mid(nezdata, i + 1)
                i = InStr(nezdata, ",")
                If i = 0 Then
                    errorcode = 20
                    Exit Sub
                End If
                Dim vdata As String = Left(nezdata, i - 1)
                If vdata = "" Then
                    errorcode = 20
                    Exit Sub
                End If
                currentvangle = rad_to_dms(CDbl(vdata))
                nezdata = Mid(nezdata, i + 1)
                If nezdata = "" Then
                    errorcode = 20
                    Exit Sub
                End If
                currentsloped = CDbl(nezdata)
                edmpoffset = 0

            Case "SOKKIA"
                edmpoffset = 0
                nezdata = RTrim(LTrim(nezdata))
                i = InStr(nezdata, " ")
                If i > 1 Then
                    currentsloped = Val(Left(nezdata, i - 1)) / 1000
                    If currentsloped = 0 Then
                        errorcode = 10
                    Else
                        a = Mid(nezdata, i + 1)
                        i = InStr(a, " ")
                        If i > 1 And Left(a, 1) <> "E" Then
                            currentvangle = CStr(Val(Left(a, 3) + "." + Mid(a, 4, 4)))
                            a = Mid(a, i + 1)
                            i = InStr(a, " ")
                            If i > 1 And Left(a, 1) <> "E" Then
                                currenthangle = CStr(Val(Left(a, 3) + "." + Mid(a, 4, 4)))
                            Else
                                errorcode = 1
                            End If
                        Else
                            errorcode = 1
                        End If
                    End If
                Else
                    errorcode = 1
                End If

            Case Else
        End Select

        If currentsloped = 0 Then
            errorcode = 100
        End If

    End Sub

    Sub take_shot()

        Dim d As String
        Dim errorcode As Integer
        Dim a As String = ""
        Dim bcc1 As String = ""
        Dim bcc2 As String = ""

        Select Case options(4)
            Case "TOPCON"

                Call clearcom()

                '------------------------------------------------
                ' Take measurement in sloped and angle mode
                '------------------------------------------------
                d = "Z34"
                Call edmoutput(d, errorcode)
                Call edminput(a)

                Call delay(0.5)

                '------------------------------------------------
                ' Take the actual measurement
                '------------------------------------------------
                d = "C"
                Call edmoutput(d, errorcode)
                Call edminput(a)

                Call delay(0.5)

            Case "WILD", "LEICA"
                Call clearcom()
                d = "GET/M/WI11/WI21/WI22/WI31/WI51"
                Call edmoutput(d, errorcode)

            Case "LEICA-GEO"
                Call clearcom()
                d = "%R1Q,2008:1,1"
                Call edmoutput(d, errorcode)
                Call delay(1)
                Call clearcom()
                d = "%R1Q,2108:10000,1"
                Call edmoutput(d, errorcode)

            Case "SOKKIA"
                d = Chr(17)
                Call edmoutput(d, errorcode)

            Case Else
        End Select

    End Sub

    Sub sethortangle(ByVal angle As String, ByVal deg As Integer, ByVal min As Integer, ByVal sec As Single)

        Dim a As Integer
        Dim d As String
        Dim errorcode As Integer
        Dim da As String
        Dim ma As String
        Dim sa As String
        Dim s As String = ""

        Select Case options(4)
            Case "LEICA-GEO"
                If angle = "" Then
                    Dim angle_dd As Double = CDbl(deg) + CDbl(min / 60.0) + CDbl(sec / 3600.0)
                    Dim angle_rad As Double = angle_dd * 0.0174532925199433
                    angle = CStr(angle_rad)
                Else
                    deg = CInt(angle)
                    If InStr(angle, ".") <> 0 Then
                        angle = angle + "0000"
                        a = InStr(angle, ".")
                        min = CInt(Mid(angle, a + 1, 2))
                        sec = CInt(Mid(angle, a + 3, 2))
                    Else
                        min = 0
                        sec = 0
                    End If
                    Dim angle_dd As Double = CDbl(deg) + CDbl(min / 60.0) + CDbl(sec / 3600.0)
                    Dim angle_rad As Double = angle_dd * 0.0174532925199433
                    angle = CStr(angle_rad)
                End If

            Case Else
                If angle = "" Then
                    da = Right("000" + LTrim(Str(deg)), 3)
                    ma = Right("00" + LTrim(Str(min)), 2)
                    sa = Right("00" + LTrim(Str(sec)), 2)
                    If UCase(options(4)) = "SOKKIA" Then
                        angle = da + "." + ma + sa
                    Else
                        angle = da + ma + sa
                    End If

                ElseIf InStr(angle, ".") <> 0 Then
                    a = InStr(angle, ".")
                    angle = Left(angle + "0000", a + 4)
                    If UCase(options(4)) = "SOKKIA" Then
                        angle = Right("000" + angle, 8)
                        If Val(Left(angle, 3)) >= 360 Then
                            angle = Right("000" + Trim(Str(Val(Left(angle, 3)) Mod 360)), 3) + Mid(angle, 4)
                        End If
                    Else
                        angle = Left(angle, a - 1) + Mid(angle, a + 1)
                    End If
                End If
        End Select

        Select Case options(4)
            Case "NIKON"
                d = "!HAN" + angle
                Call edmoutput(d, errorcode)
                Call delay(1)
                Call clearcom()

            Case "TOPCON"
                d = "J+" + LTrim(angle) + "d"
                Call edmoutput(d, errorcode)
                Call edminput(s)
                Call delay(1)
                Call clearcom()

            Case "WILD", "LEICA"
                d = "PUT/21...4+" + Right("000" + LTrim(angle) + "0 ", 9)
                Call edmoutput(d, errorcode)
                Call edminput(s)
                Call delay(5)

            Case "LEICA-GEO"
                d = "%R1Q,2113:" + angle
                Call edmoutput(d, errorcode)
                Call delay(1)
                Call edminput(s)

            Case "SOKKIA"
                d = "/Dc " + angle + Chr(13) + Chr(10)

                '***********************************
                ' Debug code - to help Jalil with
                ' a problem on Nov. 10, 2012
                '***********************************
                MsgBox("Angle will be set to '" + angle + "'", MsgBoxStyle.Information And MsgBoxStyle.OkOnly, "EDM-Mobile")
                '***********************************
                ' End debug code
                '***********************************

                'd$ = "Gd" + Chr(13) + Chr(10) gets angle

                Call edmoutput(d, errorcode)
                Call delay(5)
                Call clearcom()

            Case Else
        End Select

    End Sub

    Sub vhdtonez()

        Dim angle As Integer
        Dim minutes As Integer
        Dim seconds As Integer
        Dim tangle As Double
        Dim actuald As Double

        Call parseangle(CSng(currentvangle), angle, minutes, seconds)
        tangle = angle + ((minutes * 60 + seconds) / 3600)

        'Adjust the slope distance for the prism offset
        If currentprismoffset <> 0 And (currentprismoffset <> edmprismoffset) Then
            If edmprismoffset <> 0 Then
                currentsloped = currentsloped - edmprismoffset
            End If
            If currentprismoffset <> 0 Then
                currentsloped = currentsloped + currentprismoffset
            End If
        End If

        'If currentprismoffset = 0 And edmprismoffset = 0 Then
        '    currentsloped = currentsloped
        'ElseIf currentprismoffset <> 0 And edmprismoffset <> 0 Then
        '    If currentprismoffset <> edmprismoffset Then
        '        MsgBox "WARNING: Instrument prism offset (" + CStr(edmprismoffset) + ") and program prism offset (" + CStr(currentprismoffset) + ")do not match.  Program prism offset will be used.  Either set program prism offset to 0 and use total station offset or set total station to 0 and use program offsets.", vbCritical
        '    End If
        '    currentsloped = currentsloped + currentprismoffset
        'ElseIf currentprismoffset <> 0 And edmprismoffset = 0 Then
        '    currentsloped = currentsloped + currentprismoffset
        'ElseIf edmprismoffset <> 0 And currentprismoffset = 0 Then
        '    currentsloped = currentsloped + edmprismoffset
        'End If

        'MsgBox "currentprismoffset=" + CStr(currentprismoffset) + ", edmprismoffset=" + CStr(edmprismoffset) + ".", vbOKOnly

        'Calculate Z relative to total station
        currentzp = currentsloped * Math.Cos(degtorad(CSng(tangle)))

        'Calculate X and Y relative to total station
        actuald = Math.Sqrt(currentsloped ^ 2 - currentzp ^ 2)
        Call parseangle(CSng(currenthangle), angle, minutes, seconds)
        tangle = angle + ((minutes * 60 + seconds) / 3600)
        tangle = 450 - tangle

        currentxp = Math.Cos(degtorad(CSng(tangle))) * actuald
        currentyp = Math.Sin(degtorad(CSng(tangle))) * actuald

    End Sub

    Sub directoutput(ByVal s As String)

        Dim frmmain As New frmMain
        Dim a(Len(s)) As Byte
        'Dim b As Integer

        'For b = 0 To Len(s) - 1
        ' a(b) = CByte(Asc(Mid(s, b, 1)))
        ' Next

        If commcontrol.IsOpen Then commcontrol.Write(s)

        'frmMain.Comm1.Output = a
        'SendData(frmMain.Comm1.CommID, a)

    End Sub

    Sub delay(ByVal a As Single)

        Dim timeone As Double = Microsoft.VisualBasic.DateAndTime.Timer
        Do
        Loop Until Microsoft.VisualBasic.DateAndTime.Timer - timeone > a

    End Sub

    Sub clearcom()

        Dim timeone As Double = Microsoft.VisualBasic.DateAndTime.Timer

        If commcontrol.IsOpen Then
            Do Until commcontrol.BytesToRead = 0 Or Microsoft.VisualBasic.DateAndTime.Timer - timeone > 10
                commcontrol.ReadExisting()
            Loop
        End If

    End Sub

    Public Sub SendArrayData(ByVal hCommID As Long, ByVal baData() As Byte, ByVal nobytes As Integer)

        'Microsoft web site code to fix issues with normal comm port output

        'Dim i, lRet, iWrite

        'For i = 1 To nobytes
        'lRet = WriteFileL(hCommID, baData(i), 1, iWrite, 0)
        'Next

        '**** may need new code here

    End Sub

    Sub center(ByVal formname As Form)

        formname.Left = CInt(Screen.PrimaryScreen.Bounds.Width / 2 - formname.Width / 2)
        formname.Top = CInt(Screen.PrimaryScreen.Bounds.Height / 2 - formname.Height / 2)

    End Sub

    Sub write_edmce_ini(ByVal filename As String, ByRef status As Integer)

        'This routine write an ini for the program itself

        'Write general ini variables
        Dim iniclass As String = "[EDM-MOBILE]"
        Dim inidata(100, 2) As String
        inidata(1, 1) = "CFG_File"
        inidata(1, 2) = cfgfile

        Call WriteIni(filename, iniclass, inidata, status)
        Dim b As Integer
        For b = 1 To UBound(inidata, 1)
            inidata(b, 1) = ""
            inidata(b, 2) = ""
        Next b

    End Sub

    Sub setup_dbs()

        Dim tbname As String
        Dim dbname As String
        Dim status As Integer
        Dim a As Integer
        Dim tpath As String
        Dim tdbname As String

        dbname = options(1)
        If Not File.Exists(dbname) Then
            'The CFG refers to a db that doesn't exist,
            'but this could be because the CFG and DB were moved to a new path
            'So, first try the same path as the cfg.
            tpath = ""
            For a = Len(cfgfile) To 1 Step -1
                If Mid(cfgfile, a, 1) = "\" Then
                    tpath = Left(cfgfile, a)
                    Exit For
                End If
            Next a
            tdbname = dbname
            For a = Len(dbname) To 1 Step -1
                If Mid(dbname, a, 1) = "\" Then
                    tdbname = Mid(dbname, a + 1)
                    Exit For
                End If
            Next a
            If Not File.Exists(tpath + tdbname) Then
                Call MsgBox("Creating new database " & dbname, vbOKOnly, "EDM-Mobile")
                Call create_db(dbname)
            Else
                dbname = tpath + tdbname
            End If
        End If
        options(1) = dbname
        tbname = options(22)
        Call table_status(tbname, status)
        If status = -1 Then Call add_field_tb()

        Call table_status("edm_datums", status)
        If status = -1 Then Call add_datum_table()

        Call table_status("edm_poles", status)
        If status = -1 Then Call add_prism_table()

        Call table_status("edm_units", status)
        If status = -1 Then Call add_unit_table()

    End Sub

    Function hash(ByVal hashlen As Integer) As String

        Dim a As Integer

        hash = ""
        For a = 1 To hashlen
            hash = hash + Chr(CInt(Int(Rnd() * 25) + Asc("A")))
        Next a

    End Function

    Public Sub SendData(ByVal hCommID As Long, ByVal sData As String)

        'Dim lRet, i, iWrite
        'For i = 1 To Len(sData)
        'If Not in_emulation Then lRet = WriteFile(hCommID, ChrB(Asc(Mid(sData, i, 1))), 1, iWrite, 0)
        'Next
        '***

    End Sub

    Sub parseangle(ByRef hangle As Single, ByRef angle As Integer, ByRef minutes As Integer, ByRef seconds As Integer)

        Dim a As String
        Dim b As String
        Dim c As Integer

        a = CStr(hangle)
        c = InStr(a, ".")
        If c = 0 Then
            angle = CInt(Int(FormatNumber(a, 0)))
            minutes = 0
            seconds = 0
        Else
            a = a + "0000"
            b = leftstr(a, c - 1)
            angle = CInt(Int(FormatNumber(b, 0)))
            minutes = CInt(Int(FormatNumber(Mid(a, c + 1, 2), 0)))
            seconds = CInt(Int(FormatNumber(Mid(a, c + 3, 2), 0)))
        End If

    End Sub

    Sub setup_tbs()

        Dim a As Integer

        Dim sqlstring As String
        Dim myconnection As New SqlCeConnection("Datasource = " + options(1))
        Dim mynonquery As SqlCeCommand = myconnection.CreateCommand
        myconnection.Open()

        If unitlimits > 0 Then

            For a = 1 To unitlimits
                sqlstring = "INSERT INTO edm_units (unit,minx,miny,maxx,maxy,centerx,centery,radius) VALUES ("
                sqlstring = sqlstring + "'" + limitname(a) + "',"
                sqlstring = sqlstring + CStr(limits(a, 1)) + ","
                sqlstring = sqlstring + CStr(limits(a, 2)) + ","
                sqlstring = sqlstring + CStr(limits(a, 3)) + ","
                sqlstring = sqlstring + CStr(limits(a, 4)) + ","
                sqlstring = sqlstring + "0,"
                sqlstring = sqlstring + "0,"
                sqlstring = sqlstring + "0)"
                mynonquery.CommandText = sqlstring
                mynonquery.Connection = myconnection
                mynonquery.ExecuteNonQuery()
            Next a

        End If

        If npoles > 0 Then

            For a = 1 To npoles
                sqlstring = "INSERT INTO edm_poles (name,height,offset) VALUES ("
                sqlstring = sqlstring + "'" + polename(a) + "',"
                sqlstring = sqlstring + CStr(poledata(a, 1)) + ","
                sqlstring = sqlstring + CStr(poledata(a, 2)) + ")"
                mynonquery.CommandText = sqlstring
                mynonquery.Connection = myconnection
                mynonquery.ExecuteNonQuery()
            Next a

        End If

        If ndatums > 0 Then

            For a = 1 To ndatums
                sqlstring = "INSERT INTO edm_datums (name,x,y,z) VALUES ("
                sqlstring = sqlstring + "'" + datumname(a) + "',"
                sqlstring = sqlstring + CStr(datum(a, 1)) + ","
                sqlstring = sqlstring + CStr(datum(a, 2)) + ","
                sqlstring = sqlstring + CStr(datum(a, 2)) + ","
                mynonquery.CommandText = sqlstring
                mynonquery.Connection = myconnection
                mynonquery.ExecuteNonQuery()
            Next a

        End If

        myconnection.Close()

    End Sub

    
    Sub savevar(ByVal a As Integer, ByVal filename As String)

        Dim inidata(100, 2) As String
        Dim iniclass As String
        Dim wstatus As Integer

        iniclass = "[" + varlist(a) + "]"

        inidata(1, 1) = "Type"
        Select Case vtype(a)
            Case 1
                inidata(1, 2) = "Text"
            Case 2
                inidata(1, 2) = "Numeric"
            Case 3
                inidata(1, 2) = "Instrument"
            Case 4
                inidata(1, 2) = "Menu"
            Case 5
                inidata(1, 2) = "Unit"
        End Select

        inidata(2, 1) = "Prompt"
        inidata(2, 2) = varprompt(a)
        inidata(3, 1) = "Menu"
        inidata(3, 2) = vmenu(a)

        If vlen(a) <> 0 Then
            inidata(4, 1) = "Length"
            inidata(4, 2) = FormatNumber(vlen(a), 0)
        End If

        If vincr(a) = True Then
            inidata(5, 1) = "Increment"
            inidata(5, 2) = "True"
        End If

        If vcarry(a) = True Then
            inidata(6, 1) = "Carry"
            inidata(6, 2) = "True"
        End If

        If vunique(a) = True Then
            inidata(7, 1) = "Unique"
            inidata(7, 2) = "True"
        End If

        Call WriteIni(filename, iniclass, inidata, wstatus)

    End Sub

    Sub save_button(ByVal buttno As Integer, ByVal buttcaption As String, ByVal filename As String)

        Dim iniclass As String
        Dim inidata(20, 2) As String
        Dim a As Integer
        Dim b As Integer
        Dim wstatus As Integer

        iniclass = "[Button" & LTrim(CStr(buttno)) & "]"
        inidata(1, 1) = "Title"
        inidata(1, 2) = buttcaption

        b = 1
        For a = 1 To vars
            If button_values(buttno, a) <> "" Then
                b = b + 1
                inidata(b, 1) = varlist(a)
                inidata(b, 2) = button_values(buttno, a)
            End If
        Next a

        Call WriteIni(filename, iniclass, inidata, wstatus)

    End Sub

    Function dms_to_dd(ByVal degrees As Integer, ByVal minutes As Integer, ByVal seconds As Integer) As Double

        dms_to_dd = CDbl(degrees) + CDbl(minutes / 60.0) + CDbl(seconds / 3600.0)

    End Function

    Sub open_tbs()

        Dim sql As String

        '***rsprisms.Open("edm_poles", options(1), adOpenKeyset, adLockOptimistic, adCmdTableDirect)
        '***rsdatums.Open("edm_datums", options(1), adOpenKeyset, adLockOptimistic, adCmdTableDirect)
        sql = "SELECT * FROM " & options(22) & " ORDER BY Recno"
        '***rspoints.Open(sql, options(1), adOpenKeyset, adLockOptimistic, adCmdText)
        '***rsunits.Open("edm_units", options(1), adOpenKeyset, adLockOptimistic, adCmdTableDirect)
        tablesopen = True
        refresh_point_grid = True

    End Sub

    Sub getfieldlen(ByVal fname As String, ByRef flen As Integer)

        Dim a As Integer

        For a = 1 To vars
            If LCase(varlist(a)) = Microsoft.VisualBasic.LCase(fname) Then
                flen = vlen(a)
                Exit For
            End If
        Next a

    End Sub

    Sub export_ascii(ByRef rdr As SqlCeDataReader, ByVal fname As String, ByVal write_header As MsgBoxResult)

        If Not rdr.Read Then Exit Sub

        Dim pdata(rdr.FieldCount - 1) As String
        Dim a As Integer
        Dim file1 As StreamWriter = New StreamWriter(fname)
        Dim dataline As String = ""

        If write_header = MsgBoxResult.Yes Then
            dataline = ""
            For a = 0 To rdr.FieldCount - 1
                If a = 0 Then
                    dataline = Chr(34) + rdr.GetName(a) + Chr(34)
                Else
                    dataline = dataline + "," + Chr(34) + rdr.GetName(a) + Chr(34)
                End If
            Next
            file1.WriteLine(dataline)
        End If

        Do
            dataline = ""
            For a = 0 To rdr.FieldCount - 1
                If Not IsDBNull(rdr.Item(a)) Then
                    Select Case rdr.GetDataTypeName(a)
                        Case "Float"
                            If a = 0 Then
                                dataline = dataline + CStr(rdr.Item(a))
                            Else
                                dataline = dataline + "," + CStr(rdr.Item(a))
                            End If
                        Case Else
                            If a = 0 Then
                                dataline = dataline + Chr(34) + rdr.Item(a).ToString + Chr(34)
                            Else
                                dataline = dataline + "," + Chr(34) + rdr.Item(a).ToString + Chr(34)
                            End If
                    End Select
                Else
                    Select Case rdr.GetDataTypeName(a)
                        Case "Float"
                            If a = 0 Then
                                dataline = dataline + "0"
                            Else
                                dataline = dataline + ",0"
                            End If
                        Case Else
                            If a = 0 Then
                                dataline = dataline + Chr(34) + Chr(34)
                            Else
                                dataline = dataline + "," + Chr(34) + Chr(34)
                            End If
                    End Select
                End If
            Next a
            file1.WriteLine(dataline)
        Loop While rdr.Read
        file1.Close()

    End Sub

    Sub close_tbs()

        'rsdatums.Close()
        'rspoints.Close()
        'rsprisms.Close()
        'rsunits.Close()
        tablesopen = False

    End Sub


    Sub check_field_sizes()

        Dim a As Integer
        Dim b As Integer
        Dim inidata(1, 2) As String
        Dim t As String

        Dim myconnection2 As New SqlCeConnection("Datasource='" + options(1) + "'")
        myconnection2.Open()
        Dim sqltxt As SqlCeCommand = myconnection2.CreateCommand
        sqltxt.CommandText = "SELECT * FROM " + options(22)
        Dim myreader As SqlCeDataReader
        myreader = sqltxt.ExecuteReader()

        For a = 1 To vars
            For b = 0 To myreader.FieldCount - 1
                If LCase(myreader.GetName(b)) = LCase(varlist(a)) Then
                    t = myreader.GetDataTypeName(b)
                    Select Case LCase(t)
                        Case "nvarchar"
                            '****If vlen(a) <> rspoints.Fields(varlist(a)).DefinedSize Then
                            'MsgBox "Warning: " & rspoints.Fields(varlist(a)).Name & " length in database (" & rspoints.Fields(varlist(a)).DefinedSize & ") does not match length in CFG file (" & vlen(a) & ").  CFG file will be adjusted to match database.", MsgBoxStyle.Information, "EDM-Mobile"
                            'vlen(a) = rspoints.Fields(varlist(a)).DefinedSize
                            'iniclass = "[" + varlist(a) + "]"
                            'inidata(1, 1) = "Length"
                            'inidata(1, 2) = vlen(a)
                            'Call WriteIni(cfgfile, iniclass, inidata, wstatus)
                            'End If
                        Case Else
                    End Select
                End If
            Next
        Next a
        myreader.Close()
        myconnection2.Close()

    End Sub

    Sub write_ini_header(ByVal filename As String)

        'This routine writes just the EDM section of the site cfg

        Dim inidata(20, 2) As String
        Dim iniclass As String
        Dim wstatus As Integer

        If LCase(Right(filename, 4)) <> ".cfg" Then filename = filename + ".cfg"

        'Write general ini variables
        iniclass = "[EDM]"
        inidata(1, 1) = "Sitename"
        inidata(1, 2) = options(18)
        inidata(2, 1) = "Database"
        inidata(2, 2) = options(1)
        inidata(3, 1) = "PointTable"
        inidata(3, 2) = options(22)
        inidata(4, 1) = "SQID"
        inidata(4, 2) = options(12)
        inidata(5, 1) = "Unitfields"
        inidata(5, 2) = options(17)
        inidata(6, 1) = "Limitchecking"
        If use_limitchecking = True Then inidata(6, 2) = "Yes" Else inidata(6, 2) = "No"

        inidata(7, 1) = "Total Station"
        inidata(7, 2) = options(4)
        inidata(8, 1) = "COM"
        inidata(8, 2) = options(2)
        inidata(9, 1) = "Settings"
        inidata(9, 2) = options(3)
        inidata(10, 1) = "CurrentStation"
        inidata(10, 2) = currentstationname
        inidata(11, 1) = "CurrentStationX"
        inidata(11, 2) = CStr(currentstationx)
        inidata(12, 1) = "CurrentStationY"
        inidata(12, 2) = CStr(currentstationy)
        inidata(13, 1) = "CurrentStationZ"
        inidata(13, 2) = CStr(currentstationz)
        inidata(14, 1) = "Referencedatum"
        inidata(14, 2) = referencedatum
        inidata(15, 1) = "Referencedatum2"
        inidata(15, 2) = referencedatum2
        inidata(16, 1) = "VerifyDatum"
        inidata(16, 2) = verifydatum
        inidata(17, 1) = "WriteLogFile"
        If WriteLogFile = True Then inidata(17, 2) = "Yes" Else inidata(17, 2) = "No"

        'Write general info for this config
        Call WriteIni(filename, iniclass, inidata, wstatus)

    End Sub

    Sub repair_compact_db()

        Dim src As String = "\my documents\WII_CE.sdf"
        Dim dst As String = "\my documents\wwii-2.sdf"

        Dim engine As New System.Data.SqlServerCe.SqlCeEngine("Data source = " + src)
        engine.Compact(("Data Source = " + dst))

    End Sub

End Module
