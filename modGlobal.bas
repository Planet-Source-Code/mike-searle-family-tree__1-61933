Attribute VB_Name = "modGlobal"
Option Explicit

Public FSO As New Scripting.FileSystemObject

Public Const ONN As Boolean = True
Public Const OFF As Boolean = False

Public Enum eRelationships
    eFather
    eMother
    eHusband
    eWife
    eChild
    eNone
End Enum

Public Enum eOptions
    eTreeName = 1
    eMainIndId = 2
    eSlideShowInterval = 3
    eEmailChanges = 4
End Enum

Public Enum eViewType
    eDetails
    ePedigree
    eGallery
End Enum

Public eView As Integer

Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Public Function GetAge(dDofBirth As String, dDofDeath As String) As Integer
Dim dDOB As Date
Dim dDeath As Date

Dim iDOBdd As String
Dim iDOBmm As String
Dim iDOByyyy As String

Dim iDODdd As String
Dim iDODmm As String
Dim iDODYYYY As String

    iDOByyyy = Mid(dDofBirth, 1, 4)
    iDOBmm = Mid(dDofBirth, 5, 2)
    iDOBdd = Mid(dDofBirth, 7, 2)

    If iDOBmm = "00" Then iDOBmm = "01"
    If iDOBdd = "00" Then iDOBdd = "01"
    
    dDOB = CDate(iDOBdd & "/" & iDOBmm & "/" & iDOByyyy)
    
    If Val(dDofDeath) = 0 Then
        dDeath = Date
    Else
        iDODYYYY = Mid(dDofDeath, 1, 4)
        iDODmm = Mid(dDofDeath, 5, 2)
        iDODdd = Mid(dDofDeath, 7, 2)
        
        If iDODmm = "00" Then iDODmm = "01"
        If iDODdd = "00" Then iDODdd = "01"
        
        dDeath = CDate(iDODdd & "/" & iDODmm & "/" & iDODYYYY)
    End If
    
    GetAge = DateDiff("YYYY", CDate(dDOB), CDate(dDeath))

End Function

Public Function ValidDate(ByRef sDateText As String) As Long
'This function takes in a string and checks to see if it is a valid date format
'if so it returns the numeric representation of the date as an 8 digit number
'If the date is valid then it is reformatted as dd MMM YYYY with 'Circa' if
'in the original therefore 'circa 19/04/1930' becomes 'Circa 19 Apr 1930'
Dim sRem As String
Dim iMonth As Integer
Dim iYear As Integer
Dim sCirca As String

    sRem = sDateText

'Two valid Circa formats are Ca and Circa

    If UCase(Mid(sRem, 1, 5)) = "CIRCA" Then
        sRem = Trim(Mid(sRem, 6, 14))
        sCirca = "Circa "
    End If
    
    If UCase(Mid(sRem, 1, 2)) = "CA" Then
        sRem = Trim(Mid(sRem, 3, 20))
        sCirca = "Circa "
    End If
    
    If UCase(Mid(sRem, 1, 6)) = "BEFORE" Then
        sRem = Trim(Mid(sRem, 7, 20))
        sCirca = "Before "
    End If
    
    If UCase(Mid(sRem, 1, 3)) = "BEF" Then
        sRem = Trim(Mid(sRem, 4, 20))
        sCirca = "Before "
    End If
    
    If UCase(Mid(sRem, 1, 5)) = "AFTER" Then
        sRem = Trim(Mid(sRem, 6, 20))
        sCirca = "After "
    End If
    
    If UCase(Mid(sRem, 1, 3)) = "AFT" Then
        sRem = Trim(Mid(sRem, 4, 20))
        sCirca = "After "
    End If
    
'If the remainder is recognised as a date then the job is done.
    If Len(sRem) > 8 And IsDate(sRem) Then
        If sCirca <> "" Then
            sDateText = sCirca & Format(CDate(sRem), "dd MMM YYYY")
        Else
            sDateText = Format(CDate(sRem), "dd MMM YYYY")
        End If
        ValidDate = Val(Format(CDate(sRem), "YYYY")) * 10000 + Val(Format(CDate(sRem), "MM")) * 100 + Val(Format(CDate(sRem), "dd"))
        Exit Function
    End If
    
'The remainder is a pure year number - always add 'Circa'
    If IsNumeric(sRem) Then
        If sRem > 1000 And sRem < 2099 Then
            If sCirca = "" Then sCirca = "Circa "
            sDateText = sCirca & Val(sRem)
            ValidDate = Val(sRem) * 10000
            Exit Function
        Else
            ValidDate = 0 'This means - not a valid date
            Exit Function
        End If
    End If
    
    Select Case UCase(Mid(sRem, 1, 3))
        Case "JAN"
            iMonth = 1
            sRem = Trim(Mid(sRem, 4, 20))
        Case "FEB"
            iMonth = 2
            sRem = Trim(Mid(sRem, 4, 20))
        Case "MAR"
            iMonth = 3
            sRem = Trim(Mid(sRem, 4, 20))
        Case "APR"
            iMonth = 4
            sRem = Trim(Mid(sRem, 4, 20))
        Case "MAY"
            iMonth = 5
            sRem = Trim(Mid(sRem, 4, 20))
        Case "JUN"
            iMonth = 6
            sRem = Trim(Mid(sRem, 4, 20))
        Case "JUL"
            iMonth = 7
            sRem = Trim(Mid(sRem, 4, 20))
        Case "AUG"
            iMonth = 8
            sRem = Trim(Mid(sRem, 4, 20))
        Case "SEP"
            iMonth = 9
            sRem = Trim(Mid(sRem, 4, 20))
        Case "OCT"
            iMonth = 10
            sRem = Trim(Mid(sRem, 4, 20))
        Case "NOV"
            iMonth = 11
            sRem = Trim(Mid(sRem, 4, 20))
        Case "DEC"
            iMonth = 12
            sRem = Trim(Mid(sRem, 4, 20))
    End Select
    
'The remainder is a pure year number - always add 'Circa'
    If IsNumeric(sRem) Then
        If sRem > 1000 And sRem < 2099 Then
            If sCirca = "" Then sCirca = "Circa "
            sDateText = sCirca & Mid(MonthName(iMonth), 1, 3) & " " & Val(sRem)
            ValidDate = Val(sRem) * 10000 + iMonth * 100
            Exit Function
        Else
            ValidDate = 0 'This means - not a valid date
            Exit Function
        End If
    End If
    
End Function

Public Function Showerror(sErr As String, nButtons As Integer) As Integer
    
    sErr = "ERROR: The following error has occurred." & vbCrLf & vbCrLf & sErr
    
    MsgBox sErr, vbOKOnly, "ERROR"

End Function

Public Function GetFullName(lngId As Long, Optional bDOB As Boolean, Optional bMarrName As Boolean) As String
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sSurname As String
Dim sFirstName As String
Dim sDOB As String
Dim sMarriedName As String

    SQL = "SELECT " & gccINDSURNAME & ", " & gccINDFIRSTNAMES & "," & _
                        gccINDDOBTEXT & ", " & gccINDGENDER & _
                        " FROM " & gtcINDIVIDUALS & " WHERE " & _
            gccINDID & " = " & lngId
            
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    If Not RS.EOF And Not RS.BOF Then
        sSurname = Trim(RS(gccINDSURNAME))
        sFirstName = Trim(RS(gccINDFIRSTNAMES))
        If bDOB Then
            sDOB = "b. " & RS(gccINDDOBTEXT)
        End If
        If bMarrName And RS(gccINDGENDER) = "F" Then
            RS.Close
            SQL = "Select Top 1 " & gccINDSURNAME & ", " & gccSPOMARRIAGEDATEDATE & _
                " FROM " & gtcMARRIAGES & _
                " INNER JOIN " & gtcINDIVIDUALS & " ON " & _
                gtcMARRIAGES & "." & gccSPOHUSBANDID & " = " & gtcINDIVIDUALS & "." & gccINDID & _
                " WHERE " & _
                    gccSPOWIFEID & " = " & lngId & _
                " ORDER BY " & _
                    gccSPOMARRIAGEDATEDATE & " DESC "
            
            Set RS = New ADODB.Recordset
            
            RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
            
            sMarriedName = ""
            If Not RS.BOF And Not RS.EOF Then
                sMarriedName = Trim(RS(gccINDSURNAME))
            End If
            RS.Close
            If sMarriedName <> "" Then
                GetFullName = Trim(sFirstName & " " & sMarriedName & " (nee " & sSurname & ")")
            Else
                GetFullName = Trim(sFirstName & " " & sSurname)
            End If
        Else
            GetFullName = Trim(sFirstName & " " & sSurname)
        End If
        If bDOB Then
            GetFullName = GetFullName & vbCrLf & sDOB
        End If
    Else
        GetFullName = ""
    End If

End Function

Public Function GetName(lngId As Long, Optional bFirstName As Boolean = False) As String
'Returns Surname if bFirstname is false
'Returns Firstnames is bfirstname is true
Dim RS As ADODB.Recordset
Dim SQL As String

    SQL = "SELECT " & gccINDSURNAME & ", " & gccINDFIRSTNAMES & _
                        " FROM " & gtcINDIVIDUALS & " WHERE " & _
            gccINDID & " = " & lngId
            
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    If Not RS.EOF And Not RS.BOF Then
        If bFirstName Then
            GetName = Trim(RS(gccINDFIRSTNAMES))
        Else
            GetName = Trim(RS(gccINDSURNAME))
        End If
    Else
        GetName = ""
    End If

End Function



Public Function IndGender(lngId As Long) As String
Dim RS As ADODB.Recordset
Dim SQL As String

    SQL = "SELECT " & gccINDGENDER & " FROM " & gtcINDIVIDUALS & _
            " WHERE " & gccINDID & " = " & lngId
            
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    If Not RS.BOF And Not RS.EOF Then
        IndGender = RS(gccINDGENDER)
    End If
    
End Function

Public Function GetOption(lOptId As eOptions) As String
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sErr As String

    On Error GoTo ErrSub
    
    SQL = "SELECT " & gccOPTVALUE & " FROM " & gtcOPTIONS & " WHERE " & _
            gccOPTID & " = " & lOptId
    
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    If Not RS.EOF And Not RS.BOF Then
        GetOption = RS(gccOPTVALUE)
    Else
        GetOption = ""
    End If
    
Exit Function
ErrSub:
    GetOption = ""

End Function

Public Sub ShowHelpContents(Form_Hwnd As Long, HelpCommand As Long, HelpContext As Long)
Dim nRet As Integer
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, App.ProductName
    Else
        On Error Resume Next
        nRet = WinHelp(Form_Hwnd, App.HelpFile, HelpCommand, HelpContext)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub


