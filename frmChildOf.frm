VERSION 5.00
Begin VB.Form frmChildOf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Child Of"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmChildOf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optParent 
      Caption         =   "Option1"
      Height          =   285
      Index           =   3
      Left            =   180
      TabIndex        =   11
      Top             =   2100
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.OptionButton optParent 
      Caption         =   "Option1"
      Height          =   285
      Index           =   2
      Left            =   180
      TabIndex        =   10
      Top             =   1740
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.OptionButton optParent 
      Caption         =   "Option1"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   1380
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.OptionButton optParent 
      Caption         =   "Option1"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   1035
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3450
      TabIndex        =   7
      Top             =   2490
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   2490
      TabIndex        =   6
      Top             =   2490
      Width           =   885
   End
   Begin VB.TextBox txtOtherParent 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   510
      TabIndex        =   4
      Top             =   2100
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.TextBox txtOtherParent 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   510
      TabIndex        =   3
      Top             =   1740
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.TextBox txtOtherParent 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   510
      TabIndex        =   2
      Top             =   1380
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.TextBox txtOtherParent 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   510
      TabIndex        =   1
      Top             =   1020
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.CheckBox chkNone 
      Caption         =   "None of these people"
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   630
      Width           =   3375
   End
   Begin VB.Label lblChildof 
      Caption         =   "This is also a child of..."
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4125
   End
End
Attribute VB_Name = "frmChildOf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlParentId As Long

Public Function invoke(lngId As Long) As Long

    Me.Caption = "Add Child of: " & GetFullName(lngId, False)

    If GetSpouses(lngId) = False Then
        invoke = 0
    Else
        Me.Show vbModal
        invoke = mlParentId
    End If

End Function

Private Function GetSpouses(lngId As Long) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim i As Integer
Dim Itmx As ListItem
Dim sGender As String

    sGender = IndGender(lngId)

    If sGender = "M" Then
        SQL = "SELECT * FROM " & gtcMARRIAGES & _
                " WHERE " & gccSPOHUSBANDID & " = " & lngId & _
                " ORDER BY " & gccSPOMARRIAGEDATEDATE
    Else
        SQL = "SELECT * FROM " & gtcMARRIAGES & _
                " WHERE " & gccSPOWIFEID & " = " & lngId & _
                " ORDER BY " & gccSPOMARRIAGEDATEDATE
    End If
            
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    i = 0
    Do While Not RS.EOF
        txtOtherParent(i).Visible = True
        txtOtherParent(i).Visible = True
        optParent(i).Visible = True
        If sGender = "M" Then
            txtOtherParent(i) = GetFullName(RS(gccSPOWIFEID))
            txtOtherParent(i).Tag = RS(gccSPOWIFEID)
        Else
            txtOtherParent(i) = GetFullName(RS(gccSPOHUSBANDID))
            txtOtherParent(i).Tag = RS(gccSPOHUSBANDID)
        End If
        i = i + 1
        RS.MoveNext
    Loop
    
    RS.Close
    
    If txtOtherParent(0).Text = "" Then
        GetSpouses = False
    Else
        GetSpouses = True
        optParent(0).Value = True
    End If
    
End Function

Private Sub chkNone_Click()
    If chkNone.Value = vbChecked Then
        optParent(0).Enabled = False
        optParent(1).Enabled = False
        optParent(2).Enabled = False
        optParent(3).Enabled = False
    Else
        optParent(0).Enabled = True
        optParent(1).Enabled = True
        optParent(2).Enabled = True
        optParent(3).Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    mlParentId = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
    If chkNone.Value = vbChecked Then
        mlParentId = 0
    Else
        For i = 0 To 4
            If optParent(i).Value = True Then
                mlParentId = Val(txtOtherParent(i).Tag)
                Exit For
            End If
        Next i
    End If
    Unload Me
End Sub
