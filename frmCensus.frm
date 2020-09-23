VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCensus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Census"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13230
   HelpContextID   =   8
   Icon            =   "frmCensus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   13230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   345
      Left            =   8370
      TabIndex        =   26
      Top             =   6510
      Width           =   885
   End
   Begin VB.TextBox txtRef 
      Height          =   285
      Left            =   11370
      TabIndex        =   10
      Top             =   1140
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      Picture         =   "frmCensus.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Find an Existing PLU"
      Top             =   6510
      Width           =   330
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   345
      Left            =   10230
      TabIndex        =   12
      Top             =   6510
      Width           =   885
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   330
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1110
      Width           =   3045
   End
   Begin VB.TextBox txtTown 
      Height          =   285
      Left            =   11370
      TabIndex        =   8
      Top             =   750
      Width           =   1815
   End
   Begin VB.TextBox txtParlDiv 
      Height          =   285
      Left            =   9480
      TabIndex        =   7
      Top             =   750
      Width           =   1815
   End
   Begin VB.TextBox txtDistrict 
      Height          =   285
      Left            =   7590
      TabIndex        =   6
      Top             =   750
      Width           =   1815
   End
   Begin VB.TextBox txtWard 
      Height          =   285
      Left            =   5700
      TabIndex        =   5
      Top             =   750
      Width           =   1815
   End
   Begin VB.TextBox txtCountyBorough 
      Height          =   285
      Left            =   3810
      TabIndex        =   4
      Top             =   750
      Width           =   1815
   End
   Begin VB.TextBox txtEccParish 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   750
      Width           =   1815
   End
   Begin VB.TextBox txtCounty 
      Height          =   285
      Left            =   1110
      TabIndex        =   1
      Top             =   60
      Width           =   2025
   End
   Begin VB.TextBox txtCivParish 
      Height          =   285
      Left            =   30
      TabIndex        =   2
      Top             =   750
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   345
      Left            =   11340
      TabIndex        =   11
      Top             =   6510
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   345
      Left            =   12300
      TabIndex        =   14
      Top             =   6510
      Width           =   885
   End
   Begin MSFlexGridLib.MSFlexGrid grdMain 
      Height          =   4965
      Left            =   330
      TabIndex        =   13
      Top             =   1440
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   8758
      _Version        =   393216
      Rows            =   20
      ForeColorFixed  =   -2147483641
      BackColorSel    =   8454143
      ForeColorSel    =   -2147483630
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ref:"
      Height          =   195
      Left            =   11010
      TabIndex        =   25
      Top             =   1170
      Width           =   300
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
      Height          =   225
      Left            =   930
      TabIndex        =   23
      Top             =   1140
      Width           =   945
   End
   Begin VB.Label Label8 
      Caption         =   "Parliamentary Division:"
      Height          =   255
      Left            =   11400
      TabIndex        =   22
      Top             =   450
      Width           =   1785
   End
   Begin VB.Label Label7 
      Caption         =   "Parliamentary Division:"
      Height          =   255
      Left            =   9510
      TabIndex        =   21
      Top             =   450
      Width           =   1785
   End
   Begin VB.Label Label6 
      Caption         =   "Rural District:"
      Height          =   255
      Left            =   7620
      TabIndex        =   20
      Top             =   450
      Width           =   1785
   End
   Begin VB.Label Label5 
      Caption         =   "Ward of Borough:"
      Height          =   255
      Left            =   5730
      TabIndex        =   19
      Top             =   450
      Width           =   1785
   End
   Begin VB.Label Label4 
      Caption         =   "County Borough:"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   450
      Width           =   1785
   End
   Begin VB.Label Label3 
      Caption         =   "Ecclesiastical Parish:"
      Height          =   255
      Left            =   1950
      TabIndex        =   17
      Top             =   450
      Width           =   1785
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "County:"
      Height          =   255
      Left            =   90
      TabIndex        =   16
      Top             =   90
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Civil Parish:"
      Height          =   255
      Left            =   60
      TabIndex        =   15
      Top             =   450
      Width           =   1785
   End
End
Attribute VB_Name = "frmCensus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum grdCols
    col_Id      'The indid of the person
    col_idName  'The name as on the individuals file
    col_Name    'The name as recorded on the census return
    col_rel
    col_Marr
    col_AgeM
    col_AgeF
    col_Occ
    col_Emp
    col_WHome
    col_Born
    col_State
End Enum

Private mlcnhID As Long
Private mbChanged As Boolean 'Indicates data has changed
Private miYear As Long       'The year of the census
Private lRow As Long
Private lCol As Long

Public Function invoke(IndID As Long, iYear As Integer) As Boolean

    miYear = iYear
    Me.Caption = iYear & " Census for "

    Call SetupGrid
    
    Call GetCensusInfo(IndID, iYear)

    
    Me.Show vbModal
End Function

Private Sub SetupGrid()
    lRow = 0
    lCol = 0

    With grdMain
        .Clear
        .Cols = 12
        .FixedCols = 0
        .Rows = 2
        .FixedRows = 1
        
        .ColAlignment(col_Id) = flexAlignLeftCenter
        .ColAlignment(col_idName) = flexAlignLeftCenter
        .ColAlignment(col_Name) = flexAlignLeftCenter
        .ColAlignment(col_rel) = flexAlignLeftCenter
        .ColAlignment(col_Marr) = flexAlignCenterCenter
        .ColAlignment(col_AgeM) = flexAlignCenterCenter
        .ColAlignment(col_AgeF) = flexAlignCenterCenter
        .ColAlignment(col_Occ) = flexAlignLeftCenter
        .ColAlignment(col_Emp) = flexAlignLeftCenter
        .ColAlignment(col_WHome) = flexAlignLeftCenter
        .ColAlignment(col_Born) = flexAlignLeftCenter
        .ColAlignment(col_State) = flexAlignLeftCenter
        
        .ColWidth(col_Id) = 0
        .ColWidth(col_idName) = .Width * (15 / 100)
        .ColWidth(col_Name) = .Width * (15 / 100)
        .ColWidth(col_rel) = .Width * (7 / 100)
        .ColWidth(col_Marr) = .Width * (3 / 100)
        .ColWidth(col_AgeM) = .Width * (4 / 100)
        .ColWidth(col_AgeF) = .Width * (4 / 100)
        .ColWidth(col_Occ) = .Width * (14 / 100)
        .ColWidth(col_Emp) = .Width * (9 / 100)
        .ColWidth(col_WHome) = .Width * (9 / 100)
        .ColWidth(col_Born) = .Width * (12 / 100)
        .ColWidth(col_State) = .Width * (7 / 100)
        
        .TextMatrix(0, col_idName) = "Link Name to Individual"
        .TextMatrix(0, col_Name) = "Name and Surname"
        .TextMatrix(0, col_rel) = "Relation"
        .TextMatrix(0, col_Marr) = "M/S"
        .TextMatrix(0, col_AgeM) = "M"
        .TextMatrix(0, col_AgeF) = "F"
        .TextMatrix(0, col_Occ) = "Occupation"
        .TextMatrix(0, col_Emp) = "Employer"
        .TextMatrix(0, col_WHome) = "Wkg at Home"
        .TextMatrix(0, col_Born) = "Where Born"
        .TextMatrix(0, col_State) = "State"

    End With
End Sub

Private Sub cmdAdd_Click()
    grdMain.Rows = grdMain.Rows + 1
    grdMain.Col = col_Name
    grdMain.Row = grdMain.Rows - 1
    mbChanged = True
    SwitchControls (ONN)
End Sub

Private Function GetCensusInfo(lngId As Long, iYear As Integer)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim lRow As Long
Dim sErr As String

    On Error GoTo ErrSub
    
    SQL = "Select " & gccCENCNHID & " FROM " & gtcCENSUS & " LEFT JOIN " & gtcCENSUSHEADER & " ON " & _
            gtcCENSUS & "." & gccCENCNHID & " = " & gtcCENSUSHEADER & "." & gccCNHID & " WHERE " & _
            gccCNHYEAR & " = " & iYear & " AND " & _
            gccCENINDID & " = " & lngId
    
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF And RS.BOF Then
        GetCensusInfo = False
        grdMain.Col = 2
        mlcnhID = 0
        Exit Function
    Else
        mlcnhID = RS(gccCENCNHID)
    End If
    RS.Close
    
    SQL = "SELECT * FROM " & gtcCENSUS & " LEFT JOIN " & gtcCENSUSHEADER & " ON " & _
            gtcCENSUS & "." & gccCENCNHID & " = " & gtcCENSUSHEADER & "." & gccCNHID & " WHERE " & _
            gccCNHID & " = " & mlcnhID
            
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    lRow = 0
    
    If Not RS.EOF And Not RS.BOF Then
        Me.Caption = Format(RS(gccCNHYEAR), "0000") & " Census for " & Format(RS(gccCNHADDRESS))
        txtCounty = Format(RS(gccCNHCOUNTY))
        txtCivParish = Format(RS(gccCNHCIVILPARISH))
        txtEccParish = Format(RS(gccCNHECCPARISH))
        txtCountyBorough = Format(RS(gccCNHCOUNTYBOROUGH))
        txtWard = Format(RS(gccCNHWARD))
        txtDistrict = Format(RS(gccCNHRURALDIST))
        txtParlDiv = Format(RS(gccCNHPARLDIV))
        txtTown = Format(RS(gccCNHTOWN))
        txtAddress = Format(RS(gccCNHADDRESS))
        Do While Not RS.EOF
            lRow = lRow + 1
            With grdMain
                If lRow > 1 Then
                    .Rows = .Rows + 1
                End If
                .TextMatrix(lRow, col_Id) = Format(RS(gccCENINDID))
                .TextMatrix(lRow, col_idName) = GetFullName(Val(RS(gccCENINDID)))
                .TextMatrix(lRow, col_Name) = Format(RS(gccCENNAME))
                .TextMatrix(lRow, col_rel) = Format(RS(gccCENRELATION))
                .TextMatrix(lRow, col_Marr) = Format(RS(gccCENMARRIED))
                .TextMatrix(lRow, col_AgeM) = Format(RS(gccCENAGEM))
                .TextMatrix(lRow, col_AgeF) = Format(RS(gccCENAGEF))
                .TextMatrix(lRow, col_Occ) = Format(RS(gccCENOCCUPATION))
                .TextMatrix(lRow, col_Emp) = Format(RS(gccCENEMPLOYER))
                .TextMatrix(lRow, col_WHome) = Format(RS(gccCENWORKINGATHOME))
                .TextMatrix(lRow, col_Born) = Format(RS(gccCENWHEREBORN))
                .TextMatrix(lRow, col_State) = Format(RS(gccCENDEAFDUMBBLIND))
            End With
            RS.MoveNext
        Loop
    End If
            
    grdMain.Col = 2
    grdMain.Row = 1
    lRow = grdMain.Row
    lCol = col_Name
    grdMain.ColSel = grdMain.Cols - 1
    txtEdit.Text = grdMain.TextMatrix(lRow, lCol)
    
    mbChanged = False
    SwitchControls (OFF)
    
Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function GetCensusInfo"
            
    Call Showerror(sErr, 0)


End Function

Private Sub cmdCancel_Click()
    If mbChanged Then
        If MsgBox("The data on this page has changed. Do you want to save it?", vbYesNo Or vbQuestion, Me.Caption) = vbYes Then
            If Not SaveCensus Then
                Exit Sub
            End If
        End If
    End If
    Unload Me
End Sub

Private Sub cmdFind_Click()
Dim lId As Long

    With grdMain
        lId = frmIndex.invoke(miYear - 110, miYear + 10, "", GetName(Val(.TextMatrix(lRow, col_Id))) & ", " & GetName(Val(.TextMatrix(lRow, col_Id)), True))
        If lId > 0 Then
            .TextMatrix(lRow, col_Id) = lId
            .TextMatrix(lRow, col_idName) = GetFullName(lId)
            txtEdit.Text = .TextMatrix(lRow, col_idName)
            mbChanged = True
            SwitchControls (ONN)
        End If
    End With

End Sub

Private Sub cmdHelp_Click()
    Call ShowHelpContents(Me.hWnd, HelpConstants.cdlHelpContext, Me.HelpContextID)
End Sub

Private Sub cmdSave_Click()
    Call SaveCensus
End Sub

Private Sub grdMain_GotFocus()
    With grdMain
        txtEdit.Top = .CellTop + .Top
        txtEdit.Left = .CellLeft + .Left
        txtEdit.Width = .CellWidth
        txtEdit.Height = .CellHeight
        txtEdit.Text = .TextMatrix(.Row, .Col)
        txtEdit.Visible = True
    End With
End Sub

Private Sub grdMain_RowColChange()
    If lRow <> 0 And lCol <> 0 Then
        grdMain.TextMatrix(lRow, lCol) = txtEdit.Text
    End If
    txtEdit.Top = grdMain.CellTop + grdMain.Top
    txtEdit.Left = grdMain.CellLeft + grdMain.Left
    txtEdit.Width = grdMain.CellWidth
    txtEdit.Height = grdMain.CellHeight
    txtEdit.Text = grdMain.TextMatrix(grdMain.Row, grdMain.Col)
    lRow = grdMain.Row
    lCol = grdMain.Col
    If lCol = col_idName Then
        txtEdit.Locked = True
        cmdFind.Top = grdMain.CellTop + grdMain.Top - ((cmdFind.Height - grdMain.CellHeight) / 2)
        cmdFind.Left = grdMain.CellLeft + grdMain.Left + grdMain.CellWidth
        cmdFind.Visible = True
    Else
        txtEdit.Locked = False
        cmdFind.Visible = False
    End If
    On Error Resume Next
    txtEdit.SetFocus
End Sub

Private Sub txtAddress_Change()
    mbChanged = True
    SwitchControls (ONN)
End Sub

Private Sub txtCivParish_Change()
    mbChanged = True
    SwitchControls (ONN)
End Sub

Private Sub txtCounty_Change()
    mbChanged = True
    SwitchControls (ONN)
End Sub

Private Sub txtCountyBorough_Change()
    mbChanged = True
    SwitchControls (ONN)
End Sub

Private Sub txtDistrict_Change()
    mbChanged = True
    SwitchControls (ONN)
End Sub

Private Sub txtEccParish_Change()
    mbChanged = True
    SwitchControls (ONN)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    With grdMain
        Select Case KeyCode
            Case vbKeyRight
                If txtEdit.SelStart >= Len(txtEdit) Then
                    If Shift = 0 Then
                        .TextMatrix(.Row, .Col) = txtEdit.Text
                        If .Col = col_State Then
                            If .Row < .Rows - 1 Then
                                .Row = .Row + 1
                            .Col = col_idName
                            End If
                        Else
                            .Col = .Col + 1
                        End If
                        KeyCode = 0
                    End If
                End If
            Case vbKeyLeft
                If txtEdit.SelStart = 0 Then
                    If Shift = 0 Then
                        If .Col = col_idName Then
                            If .Row > 1 Then
                                .Row = .Row - 1
                                .Col = col_State
                            End If
                        Else
                            .Col = .Col - 1
                        End If
                        KeyCode = 0
                    End If
                End If
            Case vbKeyDown
                .TextMatrix(.Row, .Col) = txtEdit.Text
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                End If
                KeyCode = 0
            Case vbKeyUp
                .TextMatrix(.Row, .Col) = txtEdit.Text
                If .Row > 1 Then
                    .Row = .Row - 1
                End If
                KeyCode = 0
        End Select
        .TextMatrix(.Row, .Col) = txtEdit.Text
    End With
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    mbChanged = True
    SwitchControls (ONN)
End Sub

Private Sub SwitchControls(bState As Boolean)
    cmdSave.Enabled = bState
End Sub

Private Function SaveCensus() As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sErr As String
Dim idx As Integer

    On Error GoTo ErrSub:

    If ValidDetails Then
        SQL = "Select * from " & gtcCENSUSHEADER & " WHERE "
        If mlcnhID = 0 Then
            SQL = SQL & "1 = 0" 'Blank recordset
        Else
            SQL = SQL & gccCNHID & " = " & mlcnhID
        End If

        Set RS = New ADODB.Recordset
        
        RS.Open SQL, gApp.cn, adOpenKeyset, adLockOptimistic
        
        If mlcnhID <> 0 Then
            If RS.EOF And RS.BOF Then   'If the original has disappeared then create a new one!
                RS.AddNew
            End If
        Else
            RS.AddNew
        End If
        
        RS(gccCNHADDRESS) = Trim(txtAddress)
        RS(gccCNHCIVILPARISH) = Trim(txtCivParish)
        RS(gccCNHCOUNTY) = Trim(txtCounty)
        RS(gccCNHCOUNTYBOROUGH) = Trim(txtCountyBorough)
        RS(gccCNHECCPARISH) = Trim(txtEccParish)
        RS(gccCNHPARLDIV) = Trim(txtParlDiv)
        RS(gccCNHREF) = Trim(txtRef)
        RS(gccCNHRURALDIST) = Trim(txtDistrict)
        RS(gccCNHTOWN) = Trim(txtTown)
        RS(gccCNHWARD) = Trim(txtWard)
        RS(gccCNHYEAR) = miYear
        
        RS.Update
        mlcnhID = RS(gccCNHID)
        
        RS.Close
        
        SQL = "DELETE FROM " & gtcCENSUS & " WHERE " & _
                gccCENCNHID & " = " & mlcnhID
                
        gApp.cn.Execute SQL
        
        SQL = "SELECT * FROM " & gtcCENSUS & " WHERE 1 = 0"
        
        Set RS = New ADODB.Recordset
        
        RS.Open SQL, gApp.cn, adOpenKeyset, adLockOptimistic
        
        For idx = 1 To grdMain.Rows - 1
            With grdMain
                If Trim(.TextMatrix(idx, col_Name)) <> "" Then
                    RS.AddNew
                    RS(gccCENCNHID) = mlcnhID
                    RS(gccCENNAME) = Trim(.TextMatrix(idx, col_Name))
                    RS(gccCENINDID) = Val(.TextMatrix(idx, col_Id))
                    RS(gccCENRELATION) = Trim(.TextMatrix(idx, col_rel))
                    RS(gccCENMARRIED) = Trim(.TextMatrix(idx, col_Marr))
                    RS(gccCENAGEM) = Trim(.TextMatrix(idx, col_AgeM))
                    RS(gccCENAGEF) = Trim(.TextMatrix(idx, col_AgeF))
                    RS(gccCENOCCUPATION) = Trim(.TextMatrix(idx, col_Occ))
                    RS(gccCENEMPLOYER) = Trim(.TextMatrix(idx, col_Emp))
                    RS(gccCENWORKINGATHOME) = Trim(.TextMatrix(idx, col_WHome))
                    RS(gccCENWHEREBORN) = Trim(.TextMatrix(idx, col_Born))
                    RS(gccCENDEAFDUMBBLIND) = Trim(.TextMatrix(idx, col_State))
                    RS.Update

                End If
            End With
        Next idx
        RS.Close
        mbChanged = False
        SwitchControls (OFF)
    End If

Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function GetCensusInfo"
            
    Call Showerror(sErr, 0)
    
End Function

Private Function ValidDetails() As Boolean
Dim sMess As String
Dim idx As Integer

    sMess = "You must specify as least one individual link."
    For idx = 1 To grdMain.Rows - 1
        If Val(grdMain.TextMatrix(idx, col_Id)) <> 0 Then
            sMess = ""
            Exit For
        End If
    Next idx

    If sMess = "" Then
        ValidDetails = True
    Else
        MsgBox "You cannot save this data because of the following errors..." & vbCrLf & vbCrLf & sMess, vbOKOnly Or vbCritical, Me.Caption
    End If

End Function

Private Sub txtParlDiv_Change()
    mbChanged = True
    SwitchControls (ONN)
End Sub

Private Sub txtRef_Change()
    mbChanged = True
    SwitchControls (ONN)
End Sub

Private Sub txtTown_Change()
    mbChanged = True
    SwitchControls (ONN)
End Sub

Private Sub txtWard_Change()
    mbChanged = True
    SwitchControls (ONN)
End Sub
