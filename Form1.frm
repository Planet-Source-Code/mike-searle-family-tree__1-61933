VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   465
      Left            =   540
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7050
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   5070
      TabIndex        =   1
      Text            =   "Bill Pitts"
      Top             =   7110
      Width           =   2355
   End
   Begin VB.PictureBox Picture1 
      Height          =   6825
      Left            =   570
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   6765
      ScaleWidth      =   9105
      TabIndex        =   0
      Top             =   180
      Width           =   9165
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Index           =   0
         Left            =   2100
         Top             =   2700
         Visible         =   0   'False
         Width           =   1275
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private X1 As Long
Private X2 As Long
Private Y1 As Long
Private Y2 As Long

Dim CPhoto As New colPhotoNames

Private Sub Check1_Click()
Dim i As Integer

    Image1(0).Visible = False
    For i = 1 To Image1.Count - 1
        Unload Image1(i)
    Next i
    
    If Check1.Value = vbChecked Then
        If CPhoto.Count > 0 Then
            Image1(0).Left = CPhoto(1).X1
            Image1(0).Top = CPhoto(1).Y1
            Image1(0).Width = CPhoto(1).X2 - CPhoto(1).X1
            Image1(0).Height = CPhoto(1).Y2 - CPhoto(1).Y1
            Image1(0).Visible = True
            Image1(0).ToolTipText = CPhoto(1).Note
        End If
            
        For i = 1 To CPhoto.Count - 1
            Load Image1(i)
            Image1(i).Left = CPhoto(i + 1).X1
            Image1(i).Top = CPhoto(i + 1).Y1
            Image1(i).Width = CPhoto(i + 1).X2 - CPhoto(i + 1).X1
            Image1(i).Height = CPhoto(i + 1).Y2 - CPhoto(i + 1).Y1
            Image1(i).Visible = True
            Image1(i).ToolTipText = CPhoto(i + 1).Note
        Next i
    End If
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 1 Then
        Y1 = Y
        X1 = X
        Image1(0).Left = X1
        Image1(0).Top = Y1
        Image1(0).Width = 0
        Image1(0).Height = 0
        Image1(0).Visible = True
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Dim bfound As Boolean

    If Shift = 1 Then
        If Button = 1 Then
            If X > X1 Then Image1(0).Width = X - X1
            If Y > Y1 Then Image1(0).Height = Y - Y1
        End If
    Else
        With CPhoto
            'Picture1.ToolTipText = ""
            For i = 1 To .Count
                If X >= .Item(i).X1 And X <= .Item(i).X2 And Y >= .Item(i).Y1 And Y <= .Item(i).Y2 Then
                    Picture1.ToolTipText = .Item(i).Note
                    Text1.Text = .Item(i).Note
                    bfound = True
                    Exit For
                End If
            Next i
        End With
        If Not bfound Then
            Picture1.ToolTipText = ""
'            Text1.Text = ""
        End If
    End If
    
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 1 Then
        Image1(0).Visible = False
        X2 = X
        Y2 = Y
        Debug.Print X1 & " " & X2 & " " & Y1 & " " & Y2 & " " & Text1.Text
        Call CPhoto.Add(X1, Y1, X2, Y2, Text1.Text)
    End If
End Sub
