VERSION 5.00
Begin VB.Form frmGetCaption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Caption"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "frmGetCaption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5115
   StartUpPosition =   1  'CenterOwner
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
      Height          =   315
      Left            =   4680
      Picture         =   "frmGetCaption.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Find an Existing PLU"
      Top             =   180
      Width           =   330
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4140
      TabIndex        =   2
      Top             =   1170
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   345
      Left            =   3180
      TabIndex        =   1
      Top             =   1170
      Width           =   885
   End
   Begin VB.TextBox txtCaption 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4425
   End
   Begin VB.Label lblInd 
      Height          =   255
      Left            =   210
      TabIndex        =   4
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmGetCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msName As String
Private mIndId As Long

Public Function invoke(ByRef lngId As Long, Optional sOldName As String = "") As String

    txtCaption = sOldName
    mIndId = lngId
    If mIndId <> 0 Then
        lblInd = "Linked to individual Id No " & mIndId
    Else
        lblInd = "Not linked to an individual"
    End If
    Me.Show vbModal
    
    lngId = mIndId
    invoke = msName
End Function

Private Sub cmdCancel_Click()
    msName = ""
    mIndId = 0
    Unload Me
End Sub

Private Sub cmdFind_Click()
    mIndId = frmIndex.invoke(1000, 2999, "", "")
    If mIndId > 0 Then
        msName = GetFullName(mIndId, False, True)
        txtCaption = msName
    End If
End Sub

Private Sub cmdOK_Click()
    msName = txtCaption
    Unload Me
End Sub
