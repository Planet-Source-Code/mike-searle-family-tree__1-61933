VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Person Index"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   HelpContextID   =   6
   Icon            =   "frmIndex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   345
      Left            =   2670
      TabIndex        =   15
      Top             =   7980
      Width           =   885
   End
   Begin VB.OptionButton optGender 
      Caption         =   "Both"
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   14
      Top             =   990
      Width           =   885
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Again"
      Default         =   -1  'True
      Height          =   315
      Left            =   4860
      TabIndex        =   13
      Top             =   870
      Width           =   1335
   End
   Begin VB.OptionButton optGender 
      Caption         =   "Male"
      Height          =   195
      Index           =   0
      Left            =   1485
      TabIndex        =   11
      Top             =   990
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton optGender 
      Caption         =   "Female"
      Height          =   195
      Index           =   1
      Left            =   2295
      TabIndex        =   10
      Top             =   990
      Width           =   885
   End
   Begin VB.TextBox txtDOBTo 
      Height          =   285
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   8
      Top             =   660
      Width           =   915
   End
   Begin VB.TextBox txtDOBFrom 
      Height          =   285
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   6
      Top             =   360
      Width           =   915
   End
   Begin VB.TextBox txtSurname 
      Height          =   285
      Left            =   1470
      TabIndex        =   4
      Top             =   60
      Width           =   2415
   End
   Begin VB.CommandButton cmdNewPerson 
      Caption         =   "New Person"
      Height          =   345
      Left            =   30
      TabIndex        =   3
      Top             =   7980
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   4380
      TabIndex        =   2
      Top             =   7980
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5340
      TabIndex        =   1
      Top             =   7980
      Width           =   885
   End
   Begin MSComctlLib.ListView lvIndex 
      Height          =   6675
      Left            =   0
      TabIndex        =   0
      Top             =   1260
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   11774
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Gender:"
      Height          =   195
      Left            =   780
      TabIndex        =   12
      Top             =   990
      Width           =   570
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Year of Birth To:"
      Height          =   195
      Left            =   225
      TabIndex        =   9
      Top             =   690
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Year of Birth From:"
      Height          =   195
      Left            =   75
      TabIndex        =   7
      Top             =   390
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Surnames: "
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   90
      Width           =   1185
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlNewId As Long

Public Function invoke(dDOBFrom As Long, dDOBTo As Long, sGender As String, Optional sName As String = "") As Long

    txtDOBFrom = Format(dDOBFrom, "0000")
    txtDOBTo = Format(dDOBTo, "0000")
    mlNewId = -1
    If Trim(sName) <> "," Then
        txtSurname = sName
    End If
    
    Select Case sGender
        Case "M"
            optGender(0).Value = True
            optGender(0).Enabled = False
            optGender(1).Enabled = False
            optGender(2).Enabled = False
        Case "F"
            optGender(1).Value = True
            optGender(0).Enabled = False
            optGender(1).Enabled = False
            optGender(2).Enabled = False
        Case Else
            optGender(2).Value = True
    End Select
    
    Call SearchIndex
    
    Me.Show vbModal
    
    invoke = mlNewId
    
End Function

Private Function SearchIndex()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim idx As Integer
Dim Itmx As ListItem
Dim sErr As String

Dim lDateFrom As Long
Dim lDateTo As Long

Dim sNames() As String

    On Error GoTo ErrSub
    
    lDateFrom = (Val(txtDOBFrom) * 10000) + 101
    lDateTo = (Val(txtDOBTo) * 10000) + 1231

    If Trim(txtSurname) = "" Then
        ReDim sNames(1)
        sNames(0) = ""
    Else
        sNames = Split(Trim(txtSurname), ",")
    End If

    With lvIndex
        .ColumnHeaders.Clear
        .ListItems.Clear

        Call .ColumnHeaders.Add(1, "Name", "Name", 2800)
        .ColumnHeaders(1).Tag = lvwAscending
        Call .ColumnHeaders.Add(2, "Date of Birth", "Date of Birth", 1500)
        .ColumnHeaders(2).Tag = lvwAscending
        Call .ColumnHeaders.Add(3, "Place of Birth", "Place of Birth", 2500)
        .ColumnHeaders(3).Tag = lvwAscending
        .View = lvwReport
    End With


    SQL = "SELECT " & gccINDID & ", " & gccINDSURNAME & ", " & gccINDFIRSTNAMES & ", " & _
                gccINDDOBTEXT & ", " & gccINDPLACEOFBIRTH & ", " & _
                gccINDBAPTDATETEXT & ", " & gccINDBAPTCHURCH & _
                " FROM " & _
                gtcINDIVIDUALS & " WHERE "
                
    SQL = SQL & "((" & _
                gccINDDOBDATE & " >= " & lDateFrom & " OR " & _
                gccINDDOBDATE & " = 0) AND (" & _
                gccINDDOBDATE & " <= " & lDateTo & " OR " & _
                gccINDDOBDATE & " = 0)) "
'                OR ((" & _
'                gccINDBAPTDATEDATE & " >= " & lDateFrom & " OR " & _
'                gccINDBAPTDATEDATE & " = 0) AND (" & _
'                gccINDBAPTDATEDATE & " <= " & lDateTo & " OR " & _
'                gccINDBAPTDATEDATE & " = 0))) "
    
    If Trim(txtSurname) <> "" Then
        SQL = SQL & " AND " & gccINDSURNAME & " LIKE '" & Trim(sNames(0)) & "%'"
    End If
    
    If UBound(sNames) > 0 Then
        SQL = SQL & " AND " & gccINDFIRSTNAMES & " Like '" & Trim(sNames(1)) & "%'"
    End If
    
    If optGender(0).Value = True Then SQL = SQL & " AND " & gccINDGENDER & " = 'M'"
    If optGender(1).Value = True Then SQL = SQL & " AND " & gccINDGENDER & " = 'F'"
    
    
    SQL = SQL & " ORDER BY " & gccINDDOBDATE

    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not RS.EOF
        idx = idx + 1
        Set Itmx = lvIndex.ListItems.Add(idx, RS(gccINDID) & "X", UCase(Trim(RS(gccINDSURNAME))) & ", " & Trim(RS(gccINDFIRSTNAMES)))
        If Not IsNull(RS(gccINDDOBTEXT)) Then
            Itmx.SubItems(1) = RS(gccINDDOBTEXT)
        Else
            If Not IsNull(RS(gccINDBAPTDATETEXT)) Then
                Itmx.SubItems(1) = RS(gccINDBAPTDATETEXT)
            End If
        End If
        If Not IsNull(RS(gccINDPLACEOFBIRTH)) Then
            Itmx.SubItems(2) = RS(gccINDPLACEOFBIRTH)
        Else
            If Not IsNull(RS(gccINDBAPTCHURCH)) Then
                Itmx.SubItems(2) = RS(gccINDBAPTCHURCH)
            End If
        End If
        RS.MoveNext
    Loop
    RS.Close

Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function SearchIndex"

    Call Showerror(sErr, 0)

End Function

Private Sub cmdCancel_Click()
    mlNewId = -1
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelpContents(Me.hWnd, HelpConstants.cdlHelpContext, Me.HelpContextID)
End Sub

Private Sub cmdNewPerson_Click()
    mlNewId = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Val(lvIndex.SelectedItem.Key) > 0 Then
        mlNewId = Val(lvIndex.SelectedItem.Key)
        Unload Me
    End If
End Sub

Private Sub cmdSearch_Click()
    Call SearchIndex
End Sub

Private Sub lvIndex_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Tag = lvwAscending Then
        ColumnHeader.Tag = lvwDescending
    Else
        ColumnHeader.Tag = lvwAscending
    End If
    lvIndex.SortOrder = ColumnHeader.Tag
    lvIndex.SortKey = ColumnHeader.Index - 1
    lvIndex.Sorted = True

End Sub

Private Sub lvIndex_DblClick()
    cmdOK_Click
End Sub

Private Sub optGender_Click(Index As Integer)
    cmdSearch_Click
End Sub
