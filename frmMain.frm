VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The ... Family Tree"
   ClientHeight    =   8565
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11880
   HelpContextID   =   2
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSMAPI.MAPIMessages MAPIMes 
      Left            =   6780
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISes 
      Left            =   6150
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame fraDetails 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   7905
      Left            =   1020
      TabIndex        =   21
      Top             =   2940
      Width           =   11865
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Sources"
         Height          =   975
         Left            =   5070
         TabIndex        =   41
         Top             =   510
         Width           =   6765
         Begin VB.CommandButton cmdOther 
            Caption         =   "Other"
            Height          =   315
            Left            =   4140
            TabIndex        =   54
            Top             =   570
            Width           =   1305
         End
         Begin VB.CommandButton cmdCensus 
            Caption         =   "1901 Census"
            Height          =   315
            Index           =   6
            Left            =   4140
            TabIndex        =   53
            Top             =   240
            Width           =   1305
         End
         Begin VB.CommandButton cmdCensus 
            Caption         =   "1891 Census"
            Height          =   315
            Index           =   5
            Left            =   2790
            TabIndex        =   52
            Top             =   570
            Width           =   1305
         End
         Begin VB.CommandButton cmdCensus 
            Caption         =   "1881 Census"
            Height          =   315
            Index           =   4
            Left            =   2790
            TabIndex        =   51
            Top             =   240
            Width           =   1305
         End
         Begin VB.CommandButton cmdCensus 
            Caption         =   "1871 Census"
            Height          =   315
            Index           =   3
            Left            =   1440
            TabIndex        =   50
            Top             =   570
            Width           =   1305
         End
         Begin VB.CommandButton cmdCensus 
            Caption         =   "1861 Census"
            Height          =   315
            Index           =   2
            Left            =   1440
            TabIndex        =   49
            Top             =   240
            Width           =   1305
         End
         Begin VB.CommandButton cmdCensus 
            Caption         =   "1851 Census"
            Height          =   315
            Index           =   1
            Left            =   90
            TabIndex        =   48
            Top             =   570
            Width           =   1305
         End
         Begin VB.CommandButton cmdCensus 
            Caption         =   "1841 Census"
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   42
            Top             =   240
            Width           =   1305
         End
      End
      Begin VB.Frame fraParents 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Parents"
         Height          =   975
         Left            =   0
         TabIndex        =   36
         Top             =   510
         Width           =   5025
         Begin VB.Label lblMother 
            BackColor       =   &H00E7DEFE&
            Caption         =   "Mothers Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   810
            TabIndex        =   104
            Top             =   570
            Width           =   3645
         End
         Begin VB.Label lblFather 
            BackColor       =   &H00FEFFD7&
            Caption         =   "Fathers Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   810
            TabIndex        =   103
            Top             =   240
            Width           =   3645
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Mother:"
            Height          =   195
            Left            =   180
            TabIndex        =   38
            Top             =   600
            Width           =   540
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Father:"
            Height          =   195
            Left            =   240
            TabIndex        =   37
            Top             =   270
            Width           =   495
         End
      End
      Begin VB.Frame fraChildren 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Children"
         Height          =   4815
         Left            =   5070
         TabIndex        =   35
         Top             =   3090
         Width           =   6765
         Begin VB.Label lblChildDOB 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   3870
            TabIndex        =   102
            Top             =   210
            Width           =   2115
         End
         Begin VB.Label lblChild 
            Caption         =   "Childrens Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   101
            Top             =   210
            Width           =   3645
         End
      End
      Begin VB.Frame fraAddress 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Current or last known address"
         Height          =   2565
         Left            =   0
         TabIndex        =   34
         Top             =   5340
         Width           =   5025
         Begin VB.TextBox txtPostcode 
            Height          =   285
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   15
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox txtCounty 
            Height          =   285
            Left            =   1350
            MaxLength       =   20
            TabIndex        =   14
            Top             =   1140
            Width           =   2295
         End
         Begin VB.TextBox txtTown 
            Height          =   285
            Left            =   1350
            MaxLength       =   30
            TabIndex        =   13
            Top             =   840
            Width           =   3585
         End
         Begin VB.TextBox txtEmail 
            Height          =   285
            Left            =   1350
            MaxLength       =   60
            TabIndex        =   17
            Top             =   2040
            Width           =   3585
         End
         Begin VB.TextBox txtPhone 
            Height          =   285
            Left            =   1350
            MaxLength       =   30
            TabIndex        =   16
            Top             =   1740
            Width           =   3585
         End
         Begin VB.TextBox txtAddress 
            Height          =   525
            Left            =   1350
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   270
            Width           =   3585
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Postcode:"
            Height          =   195
            Left            =   480
            TabIndex        =   57
            Top             =   1500
            Width           =   720
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "County:"
            Height          =   195
            Left            =   660
            TabIndex        =   56
            Top             =   1200
            Width           =   540
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Town:"
            Height          =   195
            Left            =   750
            TabIndex        =   55
            Top             =   900
            Width           =   450
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            Height          =   195
            Left            =   720
            TabIndex        =   44
            Top             =   2100
            Width           =   420
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone:"
            Height          =   195
            Left            =   645
            TabIndex        =   43
            Top             =   1800
            Width           =   510
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   540
            TabIndex        =   39
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.Frame fraSpouses 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Spouse(s)"
         Height          =   1575
         Left            =   5070
         TabIndex        =   33
         Top             =   1470
         Width           =   6765
         Begin VB.TextBox txtMarriageDate 
            Height          =   285
            Index           =   0
            Left            =   3840
            MaxLength       =   20
            TabIndex        =   18
            Top             =   188
            Width           =   1635
         End
         Begin VB.CommandButton cmdViewMarriageCert 
            Caption         =   "Marriage Cert"
            Height          =   315
            Index           =   3
            Left            =   5550
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   1170
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CommandButton cmdViewMarriageCert 
            Caption         =   "Marriage Cert"
            Height          =   315
            Index           =   2
            Left            =   5550
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CommandButton cmdViewMarriageCert 
            Caption         =   "Marriage Cert"
            Height          =   315
            Index           =   1
            Left            =   5550
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   510
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CommandButton cmdViewMarriageCert 
            Caption         =   "Marriage Cert"
            Height          =   315
            Index           =   0
            Left            =   5550
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lblSpouse 
            Caption         =   "Spouses Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   270
            TabIndex        =   105
            Top             =   210
            Width           =   3495
         End
         Begin VB.Label lblSpouseNum 
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   96
            Top             =   240
            Width           =   135
         End
      End
      Begin VB.Frame fraIndividual 
         BackColor       =   &H00C0FFFF&
         Height          =   3825
         Left            =   0
         TabIndex        =   22
         Top             =   1470
         Width           =   5025
         Begin VB.TextBox txtSurname 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1335
            MaxLength       =   30
            TabIndex        =   106
            Top             =   240
            Width           =   3585
         End
         Begin VB.CommandButton cmdViewDCert 
            Caption         =   "Death Cert"
            Height          =   315
            Left            =   3900
            TabIndex        =   46
            Top             =   2790
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton cmdViewBCert 
            Caption         =   "Birth Cert"
            Height          =   315
            Left            =   3900
            TabIndex        =   45
            Top             =   1140
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CheckBox chkAdopted 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Adopted"
            Height          =   195
            Left            =   3240
            TabIndex        =   3
            Top             =   900
            Width           =   1335
         End
         Begin VB.TextBox txtBuriedAt 
            Height          =   285
            Left            =   1335
            MaxLength       =   50
            TabIndex        =   11
            Top             =   3420
            Width           =   3585
         End
         Begin VB.TextBox txtPlaceofDeath 
            Height          =   285
            Left            =   1335
            MaxLength       =   50
            TabIndex        =   10
            Top             =   3120
            Width           =   3585
         End
         Begin VB.TextBox txtDateofDeath 
            Height          =   285
            Left            =   1335
            MaxLength       =   20
            TabIndex        =   9
            Top             =   2820
            Width           =   1725
         End
         Begin VB.CheckBox chkDeceased 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Deceased:"
            Height          =   225
            Left            =   405
            TabIndex        =   8
            Top             =   2580
            Width           =   1125
         End
         Begin VB.TextBox txtBaptChurch 
            Height          =   285
            Left            =   1335
            MaxLength       =   50
            TabIndex        =   7
            Top             =   2160
            Width           =   3585
         End
         Begin VB.TextBox txtPlaceOB 
            Height          =   285
            Left            =   1335
            MaxLength       =   50
            TabIndex        =   5
            Top             =   1470
            Width           =   3585
         End
         Begin VB.TextBox txtBaptised 
            Height          =   285
            Left            =   1335
            MaxLength       =   20
            TabIndex        =   6
            Top             =   1860
            Width           =   1725
         End
         Begin VB.TextBox txtDOB 
            Height          =   285
            Left            =   1335
            MaxLength       =   20
            TabIndex        =   4
            Top             =   1170
            Width           =   1725
         End
         Begin VB.OptionButton optGender 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Female"
            Height          =   195
            Index           =   1
            Left            =   2175
            TabIndex        =   2
            Top             =   900
            Width           =   885
         End
         Begin VB.OptionButton optGender 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Male"
            Height          =   195
            Index           =   0
            Left            =   1365
            TabIndex        =   1
            Top             =   900
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.TextBox txtFirstNames 
            Height          =   285
            Left            =   1335
            MaxLength       =   30
            TabIndex        =   0
            Top             =   540
            Width           =   3585
         End
         Begin VB.Label lblAge 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Age (99)"
            Height          =   255
            Left            =   3120
            TabIndex        =   97
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Buried at:"
            Height          =   195
            Left            =   570
            TabIndex        =   32
            Top             =   3450
            Width           =   675
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Place of Death:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   3150
            Width           =   1110
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Died:"
            Height          =   195
            Left            =   450
            TabIndex        =   30
            Top             =   2850
            Width           =   765
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "In Church:"
            Height          =   195
            Left            =   495
            TabIndex        =   29
            Top             =   2190
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Place of Birth:"
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   1500
            Width           =   990
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Baptised on:"
            Height          =   195
            Left            =   345
            TabIndex        =   27
            Top             =   1890
            Width           =   885
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth:"
            Height          =   195
            Left            =   300
            TabIndex        =   26
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender:"
            Height          =   195
            Left            =   660
            TabIndex        =   25
            Top             =   900
            Width           =   570
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name(s):"
            Height          =   195
            Left            =   270
            TabIndex        =   24
            Top             =   570
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Surname:"
            Height          =   195
            Left            =   555
            TabIndex        =   23
            Top             =   270
            Width           =   675
         End
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Click on a persons name to show details for that person"
         Height          =   495
         Left            =   8880
         TabIndex        =   108
         Top             =   60
         Width           =   2055
      End
      Begin VB.Label lblFullName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3090
         TabIndex        =   61
         Top             =   90
         Width           =   5385
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Details for: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   40
         Top             =   90
         Width           =   2925
      End
   End
   Begin VB.Frame fraOther 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7875
      Left            =   750
      TabIndex        =   89
      Top             =   1440
      Width           =   11865
      Begin MSComDlg.CommonDialog cdgAddPic 
         Left            =   2040
         Top             =   990
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11460
         TabIndex        =   100
         Top             =   540
         Width           =   345
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   99
         Top             =   540
         Width           =   345
      End
      Begin VB.CommandButton cmdAddPicture 
         Caption         =   "Add Picture"
         Height          =   315
         Left            =   450
         TabIndex        =   98
         Top             =   540
         Width           =   1035
      End
      Begin VB.PictureBox picGallery 
         AutoRedraw      =   -1  'True
         Height          =   1305
         Index           =   0
         Left            =   60
         ScaleHeight     =   1245
         ScaleMode       =   0  'User
         ScaleWidth      =   1425
         TabIndex        =   95
         Top             =   930
         Width           =   1425
      End
      Begin RichTextLib.RichTextBox rtbMemo 
         Height          =   2385
         Left            =   60
         TabIndex        =   92
         Top             =   5430
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   4207
         _Version        =   393217
         TextRTF         =   $"frmMain.frx":058A
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Double click on a picture to see it full size."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1710
         TabIndex        =   107
         Top             =   510
         Width           =   5265
      End
      Begin VB.Label lblImgCaption 
         BorderStyle     =   1  'Fixed Single
         Height          =   675
         Index           =   0
         Left            =   60
         TabIndex        =   94
         Top             =   2310
         Width           =   1425
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Narrative"
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
         Left            =   60
         TabIndex        =   93
         Top             =   5130
         Width           =   2595
      End
      Begin VB.Label Label28 
         Caption         =   "Other Details for: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   91
         Top             =   120
         Width           =   2475
      End
      Begin VB.Label lblOtherDetailsName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2790
         TabIndex        =   90
         Top             =   120
         Width           =   5385
      End
   End
   Begin VB.Frame fraPedigree 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7875
      Left            =   120
      TabIndex        =   62
      Top             =   2160
      Width           =   11865
      Begin VB.TextBox txtInvisible 
         Height          =   285
         Left            =   390
         TabIndex        =   88
         Text            =   "Invisible"
         Top             =   1650
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSComctlLib.ListView lvSpouses 
         Height          =   1155
         Left            =   180
         TabIndex        =   80
         Top             =   3120
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtMMFather 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   9210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   77
         Top             =   6042
         Width           =   2595
      End
      Begin VB.TextBox txtMFMother 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   9210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   76
         Top             =   5070
         Width           =   2595
      End
      Begin VB.TextBox txtMFFather 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   9210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   75
         Top             =   4098
         Width           =   2595
      End
      Begin VB.TextBox txtFMMother 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   9210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   74
         Top             =   3126
         Width           =   2595
      End
      Begin VB.TextBox txtFMFather 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   9210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   73
         Top             =   2154
         Width           =   2595
      End
      Begin VB.TextBox txtFFMother 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   9210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   72
         Top             =   1182
         Width           =   2595
      End
      Begin VB.TextBox txtFFFather 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   9210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   71
         Top             =   210
         Width           =   2595
      End
      Begin VB.TextBox txtMMMother 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   9210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   70
         Top             =   7020
         Width           =   2595
      End
      Begin VB.TextBox txtMMother 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   6270
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   69
         Top             =   6570
         Width           =   2595
      End
      Begin VB.TextBox txtMFather 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   6270
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   68
         Top             =   4640
         Width           =   2595
      End
      Begin VB.TextBox txtFMother 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   6270
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   2710
         Width           =   2595
      End
      Begin VB.TextBox txtFFather 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   6270
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   66
         Top             =   780
         Width           =   2595
      End
      Begin VB.TextBox txtPedMother 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   3330
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   65
         Top             =   5640
         Width           =   2595
      End
      Begin VB.TextBox txtPedFather 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   3330
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   64
         Top             =   1740
         Width           =   2595
      End
      Begin VB.TextBox txtMainInd 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   30
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   63
         Top             =   2430
         Width           =   2835
      End
      Begin MSComctlLib.ListView lvChildren 
         Height          =   3315
         Left            =   180
         TabIndex        =   81
         Top             =   4320
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   5847
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Maternal Grandmother"
         Height          =   195
         Left            =   6270
         TabIndex        =   87
         Top             =   6330
         Width           =   1575
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Maternal Grandfather"
         Height          =   195
         Left            =   6270
         TabIndex        =   86
         Top             =   4410
         Width           =   1500
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Paternal Grandmother"
         Height          =   195
         Left            =   6270
         TabIndex        =   85
         Top             =   2490
         Width           =   1545
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Paternal Grandfather"
         Height          =   195
         Left            =   6270
         TabIndex        =   84
         Top             =   540
         Width           =   1470
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Mother"
         Height          =   195
         Left            =   3360
         TabIndex        =   83
         Top             =   5400
         Width           =   495
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Father"
         Height          =   195
         Left            =   3330
         TabIndex        =   82
         Top             =   1500
         Width           =   450
      End
      Begin VB.Label Label21 
         Caption         =   "Pedigree for: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   79
         Top             =   90
         Width           =   1935
      End
      Begin VB.Label lblPedigreeName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2340
         TabIndex        =   78
         Top             =   90
         Width           =   5385
      End
      Begin VB.Line Line28 
         X1              =   8820
         X2              =   9060
         Y1              =   6780
         Y2              =   6780
      End
      Begin VB.Line Line27 
         X1              =   8820
         X2              =   9060
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Line26 
         X1              =   8820
         X2              =   9060
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Line Line25 
         X1              =   5880
         X2              =   6150
         Y1              =   5850
         Y2              =   5850
      End
      Begin VB.Line Line24 
         X1              =   8820
         X2              =   9060
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line23 
         X1              =   5880
         X2              =   6150
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Line Line22 
         X1              =   9060
         X2              =   9240
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line21 
         X1              =   9060
         X2              =   9240
         Y1              =   6300
         Y2              =   6300
      End
      Begin VB.Line Line20 
         X1              =   9060
         X2              =   9210
         Y1              =   5370
         Y2              =   5370
      End
      Begin VB.Line Line19 
         X1              =   9060
         X2              =   9210
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Line Line18 
         X1              =   9060
         X2              =   9210
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Line17 
         X1              =   9060
         X2              =   9210
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Line Line16 
         X1              =   9060
         X2              =   9240
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line15 
         X1              =   9060
         X2              =   9210
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Line Line14 
         X1              =   6150
         X2              =   6270
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Line Line13 
         X1              =   6150
         X2              =   6270
         Y1              =   4890
         Y2              =   4890
      End
      Begin VB.Line Line12 
         X1              =   6150
         X2              =   6270
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line11 
         X1              =   6150
         X2              =   6270
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line10 
         X1              =   2880
         X2              =   3150
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Line Line9 
         X1              =   3150
         X2              =   3330
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line8 
         X1              =   3150
         X2              =   3330
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Line Line7 
         X1              =   9060
         X2              =   9060
         Y1              =   6300
         Y2              =   7320
      End
      Begin VB.Line Line6 
         X1              =   9060
         X2              =   9060
         Y1              =   4380
         Y2              =   5370
      End
      Begin VB.Line Line5 
         X1              =   9060
         X2              =   9060
         Y1              =   2430
         Y2              =   3420
      End
      Begin VB.Line Line4 
         X1              =   9060
         X2              =   9060
         Y1              =   450
         Y2              =   1440
      End
      Begin VB.Line Line3 
         X1              =   6150
         X2              =   6150
         Y1              =   4890
         Y2              =   6840
      End
      Begin VB.Line Line2 
         X1              =   6150
         X2              =   6150
         Y1              =   1020
         Y2              =   2940
      End
      Begin VB.Line Line1 
         X1              =   3150
         X2              =   3150
         Y1              =   1980
         Y2              =   5880
      End
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1429
      ButtonWidth     =   1296
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            Key             =   "Home"
            Object.ToolTipText     =   "View Prime Individual"
            ImageKey        =   "Home"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "Back"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "Forward"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Index"
            Key             =   "Index"
            ImageKey        =   "Index"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add ..."
            Key             =   "Add"
            ImageKey        =   "Add"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Father"
                  Text            =   "Father"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Mother"
                  Text            =   "Mother"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Child"
                  Text            =   "Child"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Spouse"
                  Text            =   "Spouse"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "Unrelated"
                  Text            =   "Unrelated"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Details"
            Key             =   "Details"
            ImageKey        =   "Details"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedigree"
            Key             =   "Pedigree"
            ImageKey        =   "Pedigree"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pictures"
            Key             =   "Pictures"
            ImageKey        =   "Pictures"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Email"
            Key             =   "Email"
            Object.ToolTipText     =   "Email details"
            ImageKey        =   "Email"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "Help"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Quit"
            Key             =   "Quit"
            ImageKey        =   "Quit"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   19
      Top             =   8295
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15769
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "28/07/2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:02"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   8160
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   8700
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":060C
            Key             =   "Home"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EE6
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17C0
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":209A
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24EC
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":293E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A50
            Key             =   "File"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":332A
            Key             =   "Pictures"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C04
            Key             =   "Details"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44DE
            Key             =   "Index"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4930
            Key             =   "Pedigree"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":520A
            Key             =   "Email"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAdd 
         Caption         =   "&Add..."
         Begin VB.Menu mnuFileAddFather 
            Caption         =   "Add Father"
         End
         Begin VB.Menu mnuFileAddMother 
            Caption         =   "Add Mother"
         End
         Begin VB.Menu mnuFileAddSpouse 
            Caption         =   "Add Spouse"
         End
         Begin VB.Menu mnuFileAddChild 
            Caption         =   "Add Child"
         End
         Begin VB.Menu mnuFileAddUnrelated 
            Caption         =   "Add unrelated person"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSendTo 
         Caption         =   "Send &To..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Sen&d..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "Paste &Special..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIndividual 
         Caption         =   "&Individual"
      End
      Begin VB.Menu mnuViewPedigree 
         Caption         =   "&Pedigree"
      End
      Begin VB.Menu mnuViewGallery 
         Caption         =   "&Gallery"
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "&Web Browser"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "Help &topic for this page"
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "Help &Index"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "&PoPup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopView 
         Caption         =   "View"
      End
      Begin VB.Menu mnuPopProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuPopSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopCancel 
         Caption         =   "&Cancel"
      End
   End
   Begin VB.Menu mnuGoto 
      Caption         =   "&Goto"
      Visible         =   0   'False
      Begin VB.Menu mnuGotoAdd 
         Caption         =   "&Add Individual"
      End
      Begin VB.Menu mnuGotoInd 
         Caption         =   "&View Individual"
      End
      Begin VB.Menu mnuGotoRemove 
         Caption         =   "&Remove Relationship"
      End
      Begin VB.Menu mnuGotoSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGotoCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbChanged As Boolean 'Indicates data has changed
Private mIndex As Integer
Private mNewInd As Long
Private mDelRel As Long
Private ePersonToAdd As Integer
Private bBackorFwd As Boolean

Private mlHistory() As Long
Private mnHistIdx As Integer

'Private Enum AddPerson
'    eFather
'    eMother
'    eSpouse
'    eChild
'End Enum

Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Function UpdateRelationShips(Optional lOtherParentID As Long = 0) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim i As Integer
Dim sErr As String

    On Error GoTo ErrSub

    mDelRel = 0

    'If the father tag is value 0 an existing relationship will be removed
    SQL = "Update " & gtcINDIVIDUALS & " set " & _
            gccINDFATHERID & " = " & Val(lblFather.Tag) & _
            " WHERE " & gccINDID & " = " & Val(txtSurname.Tag)
            
    gApp.cn.Execute SQL

    'If the mother tag is value 0 an existing relationship will be removed
    SQL = "Update " & gtcINDIVIDUALS & " set " & _
            gccINDMOTHERID & " = " & Val(lblMother.Tag) & _
            " WHERE " & gccINDID & " = " & Val(txtSurname.Tag)
            
    gApp.cn.Execute SQL

    For i = 0 To lblSpouse.Count - 1
        If Val(lblSpouse(i).Tag) <> 0 Then
            If optGender(0).Value = True Then
                SQL = "SELECT Count(*) as NumRecs from " & gtcMARRIAGES & " WHERE " & _
                    gccSPOHUSBANDID & " = " & Val(txtSurname.Tag) & " AND " & _
                    gccSPOWIFEID & " = " & Val(lblSpouse(i).Tag)
            Else
                SQL = "SELECT Count(*) as NumRecs from " & gtcMARRIAGES & " WHERE " & _
                    gccSPOHUSBANDID & " = " & Val(lblSpouse(i).Tag) & " AND " & _
                    gccSPOWIFEID & " = " & Val(txtSurname.Tag)
            End If
            
            Set RS = New ADODB.Recordset
            
            RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
            
            If RS("NumRecs") = 0 Then
                If optGender(0).Value = True Then
                    SQL = "INSERT INTO " & gtcMARRIAGES & " (" & _
                        gccSPOHUSBANDID & ", " & _
                        gccSPOWIFEID & ", " & _
                        gccSPOMARRIAGEDATEDATE & ", " & _
                        gccSPOMARRIAGEDATETEXT & ", " & _
                        gccSPOMARRIEDAT & ") VALUES (" & _
                    Val(txtSurname.Tag) & ", " & _
                    Val(lblSpouse(i).Tag) & ", " & _
                    ValidDate(txtMarriageDate(i)) & ", '" & _
                    txtMarriageDate(i) & "', " & _
                    "'')"
                Else
                    SQL = "INSERT INTO " & gtcMARRIAGES & " (" & _
                        gccSPOHUSBANDID & ", " & _
                        gccSPOWIFEID & ", " & _
                        gccSPOMARRIAGEDATEDATE & ", " & _
                        gccSPOMARRIAGEDATETEXT & ", " & _
                        gccSPOMARRIEDAT & ") VALUES (" & _
                    Val(lblSpouse(i).Tag) & ", " & _
                    Val(txtSurname.Tag) & ", " & _
                    ValidDate(txtMarriageDate(i)) & ", '" & _
                    txtMarriageDate(i) & "', " & _
                    "'')"
                End If
            Else
                If optGender(0).Value = True Then
                    If lblSpouse(i).Caption = "" Then
                        SQL = "DELETE  FROM " & gtcMARRIAGES & _
                                " WHERE " & _
                            gccSPOHUSBANDID & " = " & Val(txtSurname.Tag) & " AND " & _
                            gccSPOWIFEID & " = " & Val(lblSpouse(i).Tag)
                        lblSpouse(i).Tag = ""
                    Else
                        SQL = "UPDATE " & gtcMARRIAGES & " SET " & _
                            gccSPOMARRIAGEDATEDATE & " = " & ValidDate(txtMarriageDate(i)) & ", " & _
                            gccSPOMARRIAGEDATETEXT & " = '" & txtMarriageDate(i) & "' WHERE " & _
                            gccSPOHUSBANDID & " = " & Val(txtSurname.Tag) & " AND " & _
                            gccSPOWIFEID & " = " & Val(lblSpouse(i).Tag)
                    End If
                Else
                    If lblSpouse(i).Caption = "" Then
                        SQL = "DELETE  FROM " & gtcMARRIAGES & _
                                " WHERE " & _
                            gccSPOHUSBANDID & " = " & Val(lblSpouse(i).Tag) & " AND " & _
                            gccSPOWIFEID & " = " & Val(lblSpouse(i).Tag)
                        lblSpouse(i).Tag = ""
                    Else
                        SQL = "UPDATE " & gtcMARRIAGES & " SET " & _
                            gccSPOMARRIAGEDATEDATE & " = " & ValidDate(txtMarriageDate(i)) & ", " & _
                            gccSPOMARRIAGEDATETEXT & " = '" & txtMarriageDate(i) & "' WHERE " & _
                            gccSPOHUSBANDID & " = " & Val(lblSpouse(i).Tag) & " AND " & _
                            gccSPOWIFEID & " = " & Val(txtSurname.Tag)
                    End If
                End If
            End If
        
            gApp.cn.Execute SQL
        End If
    Next i
    
    For i = 0 To lblChild.Count - 1
        If Val(lblChild(i).Tag) <> 0 Then
            SQL = "Update " & gtcINDIVIDUALS & " set "
            If optGender(1).Value = True Then
                If lblChild(i).Caption = "" Then
                    SQL = SQL & gccINDMOTHERID & " = 0 "
                Else
                    SQL = SQL & gccINDMOTHERID & " = " & Val(txtSurname.Tag)
                End If
            Else
                If lblChild(i).Caption = "" Then
                    SQL = SQL & gccINDFATHERID & " = 0"
                Else
                    SQL = SQL & gccINDFATHERID & " = " & Val(txtSurname.Tag)
                End If
            End If
            SQL = SQL & " WHERE " & gccINDID & " = " & Val(Val(lblChild(i).Tag))

            gApp.cn.Execute SQL
        End If
        
        If lOtherParentID <> 0 Then
            SQL = "Update " & gtcINDIVIDUALS & " set "
            If optGender(1).Value = True Then
                If lblChild(i).Caption = "" Then
                    SQL = SQL & gccINDFATHERID & " = 0"
                Else
                    SQL = SQL & gccINDFATHERID & " = " & lOtherParentID
                End If
            Else
                If lblChild(i).Caption = "" Then
                    SQL = SQL & gccINDMOTHERID & " = 0 "
                Else
                    SQL = SQL & gccINDMOTHERID & " = " & lOtherParentID
                End If
            End If
            SQL = SQL & " WHERE " & gccINDID & " = " & Val(Val(lblChild(i).Tag))
            
            If optGender(1).Value = True Then
                SQL = SQL & " AND " & gccINDFATHERID & " = 0"
            Else
                SQL = SQL & " AND " & gccINDMOTHERID & " = 0"
            End If
            
            gApp.cn.Execute SQL

        End If
    Next i
    
    GetIndividual (Val(txtSurname.Tag))

Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function UpdateRelationShips"
            
    Call Showerror(sErr, 0)

End Function

Private Sub chkAdopted_Click()
    Call SwitchControls(ONN)
End Sub

Private Sub chkDeceased_Click()
    SwitchControls (ONN)
    If chkDeceased.Value = vbChecked Then
        txtDateofDeath.Enabled = True
        txtPlaceofDeath.Enabled = True
        txtBuriedAt.Enabled = True
    Else
        txtDateofDeath.Enabled = False
        txtPlaceofDeath.Enabled = False
        txtBuriedAt.Enabled = False
    End If
End Sub

Private Sub cmdAddPicture_Click()
Dim sFileName As String
Dim sErr As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim lNewId As Long
Dim s() As String

    On Error GoTo ErrSub

    cdgAddPic.FileName = ""
    cdgAddPic.Filter = "JPegs (*.jpg)|*.jpg|Bitmaps (*.bmp)|*.bmp"
    cdgAddPic.ShowOpen
    sFileName = cdgAddPic.FileName

    If sFileName <> "" Then
        SQL = "SELECT * FROM " & gtcIMAGES & " WHERE 1 = 0"
        
        Set RS = New ADODB.Recordset
        
        RS.Open SQL, gApp.cn, adOpenKeyset, adLockOptimistic
        
        s = Split(sFileName, "\")
        If UBound(s) > 0 Then
            sFileName = s(UBound(s))
            RS.AddNew
            RS(gccIMGNAME) = sFileName
            RS(gccIMGCAPTION) = sFileName
            RS(gccIMGDATETEXT) = "Unknown"
            RS(gccIMGDATEDATE) = 0
            RS(gccIMGNOTES) = ""
            RS.Update
            lNewId = RS(gccIMGID)
            RS.Close
            
            SQL = "SELECT * FROM " & gtcIMAGELINK & " WHERE " & _
                gccIMLIMGID & " = " & lNewId & " AND " & _
                gccIMLINDID & " = " & Val(txtSurname.Tag)
                
            Set RS = New ADODB.Recordset
            
            RS.Open SQL, gApp.cn, adOpenKeyset, adLockOptimistic
            
            If RS.EOF And RS.BOF Then
                RS.AddNew
                RS(gccIMLIMGID) = lNewId
                RS(gccIMLINDID) = Val(txtSurname.Tag)
                RS.Update
            End If
            RS.Close
            
            If picGallery.Count > 1 Then
                Call LoadNewPicBox(picGallery.Count)
            End If
            If picGallery.Count = 1 Then
                If picGallery(0).Tag <> "" Then
                    Call LoadNewPicBox(picGallery.Count)
                End If
            End If
            Call PopulatePicBox(picGallery.Count - 1, lNewId, sFileName, sFileName)
        End If
    End If
    
Exit Sub
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function cmdAddPicture_Click"
            
    Call Showerror(sErr, 0)

End Sub

Private Sub cmdCensus_Click(Index As Integer)
Dim iYear As Integer

    iYear = Index * 10
    iYear = iYear + 1841
    Call frmCensus.invoke(Val(txtSurname.Tag), iYear)
End Sub

Private Sub cmdOther_Click()
    mnuViewGallery_Click
End Sub

Private Sub Form_Load()
    'Get the primary individual
    Me.Caption = "The " & GetOption(eTreeName) & " family tree"
    fraDetails.Top = 810
    fraDetails.Left = 0
    fraPedigree.Top = 10000
    fraOther.Top = 10000
    
    ReDim Preserve mlHistory(1)
    mlHistory(0) = GetOption(eMainIndId)
    
    Call GetIndividual(GetOption(eMainIndId))
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mbChanged Then
        If MsgBox("Ok to save the changes?", vbYesNo Or vbQuestion, Me.Caption) = vbYes Then
            Call SaveIndividual
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
End Sub

Private Sub fraChildren_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HighLightChild (-1)
End Sub

Private Sub fraParents_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblFather.ForeColor = vbBlack
    lblFather.FontUnderline = False
    lblMother.ForeColor = vbBlack
    lblMother.FontUnderline = False
End Sub

Private Sub fraSpouses_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HighLightSpouse (-1)
End Sub

Private Sub lblChild_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        mNewInd = Val(lblChild(Index).Tag)
        mDelRel = Val(lblChild(Index).Tag)
        mnuGotoAdd.Visible = True
        mnuGotoInd.Visible = True
        ePersonToAdd = eChild
        PopupMenu mnuGoto
        mNewInd = 0
        mDelRel = 0
    End If
End Sub

Private Sub lblFather_Click()
    If Val(lblFather.Tag) <> 0 Then
        Call GetIndividual(Val(lblFather.Tag))
    Else
        Call mnuFileAddFather_Click
    End If
End Sub

Private Sub lblFather_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        mNewInd = Val(lblFather.Tag)
        mDelRel = mNewInd
        If mNewInd = 0 Then
            mnuGotoAdd.Visible = True
            mnuGotoInd.Visible = False
        Else
            mnuGotoAdd.Visible = False
            mnuGotoInd.Visible = True
        End If
        ePersonToAdd = eFather
        PopupMenu mnuGoto
    End If
End Sub

Private Sub lblFather_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblFather.ForeColor = vbBlue
    lblFather.FontUnderline = True
    lblMother.ForeColor = vbBlack
    lblMother.FontUnderline = False
End Sub

Private Sub lblMother_Click()
    If Val(lblMother.Tag) <> 0 Then
        Call GetIndividual(Val(lblMother.Tag))
    Else
        Call mnuFileAddMother_Click
    End If
End Sub

Private Sub lblMother_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        mNewInd = Val(lblMother.Tag)
        mDelRel = mNewInd
        If mNewInd = 0 Then
            mnuGotoAdd.Visible = True
            mnuGotoInd.Visible = False
        Else
            mnuGotoAdd.Visible = False
            mnuGotoInd.Visible = True
        End If
        ePersonToAdd = eMother
        PopupMenu mnuGoto
    End If
End Sub

Private Sub lblMother_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMother.ForeColor = vbBlue
    lblMother.FontUnderline = True
    lblFather.ForeColor = vbBlack
    lblFather.FontUnderline = False
End Sub

Private Sub lblSpouse_Click(Index As Integer)
    If Val(lblSpouse(Index).Tag) <> 0 Then
        Call GetIndividual(Val(lblSpouse(Index).Tag))
    End If
End Sub

Private Sub lblSpouse_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        mNewInd = Val(lblSpouse(Index).Tag)
        mDelRel = Val(lblSpouse(Index).Tag)
        mnuGotoAdd.Visible = True
        mnuGotoInd.Visible = True
        If optGender(0).Value = True Then
            ePersonToAdd = eWife
        Else
            ePersonToAdd = eHusband
        End If
        PopupMenu mnuGoto
        mNewInd = 0
        mDelRel = 0
    End If
End Sub

Private Sub lblSpouse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    HighLightSpouse (Index)
End Sub

Private Sub lvChildren_Click()
Dim Itmx As ListItem

    If Not lvChildren.SelectedItem Is Nothing Then
        Set Itmx = lvChildren.SelectedItem
    End If
    
    If Not Itmx Is Nothing Then
        Call GetPedigree(Val(lvChildren.SelectedItem.Key))
    End If

    On Error Resume Next
    txtInvisible.SetFocus
End Sub

Private Sub lvSpouses_Click()
Dim Itmx As ListItem

    If Not lvSpouses.SelectedItem Is Nothing Then
        Set Itmx = lvSpouses.SelectedItem
    End If
    
    If Not Itmx Is Nothing Then
        Call GetPedigree(Val(lvSpouses.SelectedItem.Key))
    End If

    On Error Resume Next
    txtInvisible.SetFocus

End Sub

Private Sub mnuFileAddChild_Click()
Dim sSurname As String
Dim sFirstNames As String
Dim lNewId As Long
Dim lDobFrom As Long
Dim lDOBTo As Long
Dim idx As Integer
Dim sGender As String
Dim lSpouseID As Long
Dim sErr As String

    On Error GoTo ErrSub

    If lblChild(0).Tag = "" Then
        idx = 0
    Else
        idx = lblChild.Count
        Load lblChild(idx)
        Load lblChildDOB(idx)
        lblChild(idx).Top = lblChild(idx - 1).Top + lblChild(idx - 1).Height
        lblChildDOB(idx).Top = lblChild(idx).Top
        lblChild(idx).Tag = ""
        lblChild(idx).Caption = ""
        lblChildDOB(idx).Caption = ""
        lblChild(idx).Visible = True
        lblChildDOB(idx).Visible = True
    End If

    lDobFrom = ValidDate(Trim(txtDOB))
    lDOBTo = ValidDate(Trim(txtDOB))
    
    lDobFrom = Int(lDobFrom / 10000) + 10
    lDOBTo = Int(lDOBTo / 10000) + 80
    
    If lDobFrom < 1000 Then lDobFrom = 1000
    If lDOBTo < 1000 Then lDOBTo = 2099
    
    If Val(lblChild(idx).Tag) = 0 Then
        If optGender(0).Value = True Then
            sSurname = txtSurname
        End If
        sFirstNames = ""
        lNewId = frmIndex.invoke(lDobFrom, lDOBTo, "")
        If lNewId = -1 Then Exit Sub
        If lNewId = 0 Then
            lSpouseID = frmChildOf.invoke(Val(txtSurname.Tag))
            lNewId = frmNewPerson.invoke(Val(txtSurname.Tag), lblFullName, eChild, sSurname, sFirstNames)
        End If
        If lNewId <> 0 Then
            lblChild(idx) = GetFullName(lNewId)
            lblChild(idx).Tag = lNewId
            Call UpdateRelationShips(lSpouseID)
            Call GetIndividual(Val(txtSurname.Tag)) 'Reload the information
        End If
    Else
        Call GetIndividual(Val(lblChild(idx).Tag))
    End If
    
Exit Sub
ErrSub:

    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function mnuFileAddChild_Click"
            
    Call Showerror(sErr, 0)


End Sub

Private Sub mnuFileAddFather_Click()
Dim sSurname As String
Dim sFirstNames As String
Dim lNewId As Long
Dim lDobFrom As Long
Dim lDOBTo As Long

    lDobFrom = ValidDate(Trim(txtDOB))
    lDOBTo = ValidDate(Trim(txtDOB))
    
    lDobFrom = Int(lDobFrom / 10000) - 100
    lDOBTo = Int(lDOBTo / 10000) - 10
    
    If lDobFrom < 1000 Then lDobFrom = 1000
    If lDOBTo < 1000 Then lDOBTo = 2099
    
    If Val(lblFather.Tag) = 0 Then
        sSurname = txtSurname
        sFirstNames = ""
        lNewId = frmIndex.invoke(lDobFrom, lDOBTo, "M")
        If lNewId = -1 Then Exit Sub
        If lNewId = 0 Then
            lNewId = frmNewPerson.invoke(Val(txtSurname.Tag), lblFullName, eFather, sSurname, sFirstNames)
        End If
        If lNewId <> 0 Then
            lblFather = GetFullName(lNewId)
            lblFather.Tag = lNewId
            Call UpdateRelationShips
            Call GetIndividual(Val(txtSurname.Tag)) 'Reload the information
        End If
    Else
        Call GetIndividual(Val(lblFather.Tag))
    End If
End Sub

Private Sub mnuFileAddMother_Click()
Dim sSurname As String
Dim sFirstNames As String
Dim lNewId As Long
Dim lDobFrom As Long
Dim lDOBTo As Long

    lDobFrom = ValidDate(Trim(txtDOB))
    lDOBTo = ValidDate(Trim(txtDOB))
    
    lDobFrom = Int(lDobFrom / 10000) - 100
    lDOBTo = Int(lDOBTo / 10000) - 10
    
    If lDobFrom < 1000 Then lDobFrom = 1000
    If lDOBTo < 1000 Then lDOBTo = 2099
    
    If Val(lblMother.Tag) = 0 Then
        sSurname = ""
        sFirstNames = ""
        lNewId = frmIndex.invoke(lDobFrom, lDOBTo, "F")
        If lNewId = -1 Then Exit Sub
        If lNewId = 0 Then
            lNewId = frmNewPerson.invoke(Val(txtSurname.Tag), lblFullName, eMother, sSurname, sFirstNames)
        End If
        If lNewId <> 0 Then
            lblMother = GetFullName(lNewId)
            lblMother.Tag = lNewId
            Call UpdateRelationShips
            Call GetIndividual(Val(txtSurname.Tag)) 'Reload the information
        End If
    Else
        Call GetIndividual(Val(lblMother.Tag))
    End If
End Sub

Private Sub mnuFileAddSpouse_Click()
Dim sSurname As String
Dim sFirstNames As String
Dim lNewId As Long
Dim lDobFrom As Long
Dim lDOBTo As Long
Dim idx As Integer
Dim sGender As String

    If lblSpouse(0).Tag = "" Then
        idx = 0
    Else
        idx = lblSpouse.Count
        Load lblSpouse(idx)
        Load txtMarriageDate(idx)
        lblSpouse(idx).Top = lblSpouse(idx - 1).Top + lblSpouse(idx - 1).Height
        txtMarriageDate(idx).Top = txtMarriageDate(idx).Top
        lblSpouse(idx).Tag = ""
        lblSpouse(idx).Caption = ""
        txtMarriageDate(idx).Text = ""
        lblSpouse(idx).Visible = True
        txtMarriageDate(idx).Visible = True
    End If

    lDobFrom = ValidDate(Trim(txtDOB))
    lDOBTo = ValidDate(Trim(txtDOB))
    
    lDobFrom = Int(lDobFrom / 10000) - 80
    lDOBTo = Int(lDOBTo / 10000) + 80
    
    If lDobFrom < 1000 Then lDobFrom = 1000
    If lDOBTo < 1000 Then lDOBTo = 2099
    
    If Val(lblSpouse(idx).Tag) = 0 Then
        sSurname = ""
        sFirstNames = ""
        'Spouse must be the oposite gender (as the law stands in 2004!)
        If optGender(0).Value = True Then
            sGender = "F"
        Else
            sGender = "M"
        End If
        lNewId = frmIndex.invoke(lDobFrom, lDOBTo, sGender)
        If lNewId = -1 Then Exit Sub
        If lNewId = 0 Then
            If sGender = "F" Then
                lNewId = frmNewPerson.invoke(Val(txtSurname.Tag), lblFullName, eWife, sSurname, sFirstNames)
            Else
                lNewId = frmNewPerson.invoke(Val(txtSurname.Tag), lblFullName, eHusband, sSurname, sFirstNames)
            End If
        End If
        If lNewId <> 0 Then
            lblSpouse(idx).Visible = True
            txtMarriageDate(idx).Visible = True
            lblSpouse(idx) = GetFullName(lNewId)
            lblSpouse(idx).Tag = lNewId
            Call UpdateRelationShips
            Call GetIndividual(Val(txtSurname.Tag)) 'Reload the information
        End If
    Else
        Call GetIndividual(Val(lblSpouse(idx).Tag))
    End If

End Sub

Private Sub mnuFileAddUnrelated_Click()
'
End Sub

Private Sub mnuFileSendTo_Click()
Dim sErr As String
Dim sAdd As String
Dim idx As Integer

    On Error GoTo ErrSub
    With MAPISes
        If Not .NewSession Then
            .DownLoadMail = False  ' Set DownLoadMail to False to prevent immediate download.
            .LogonUI = True ' Use the underlying email system's logon User ID.
            .SignOn ' Signon method.
            ' set flag to true
            .NewSession = True
        End If
    End With

    With MAPIMes
        .SessionID = MAPISes.SessionID
        .Compose
        .MsgSubject = "Some updates to the family tree."
        .MsgNoteText = "Dear " & vbCrLf & vbCrLf
        .MsgNoteText = .MsgNoteText & "Id Number:      " & txtSurname.Tag & vbCrLf
        .MsgNoteText = .MsgNoteText & "Individual:     " & lblFullName & vbCrLf
        .MsgNoteText = .MsgNoteText & "Surname:        " & txtSurname & vbCrLf
        .MsgNoteText = .MsgNoteText & "Forenames:      " & txtFirstNames & vbCrLf
        If optGender(0).Value = True Then
            .MsgNoteText = .MsgNoteText & "Gender:         " & "Male" & vbCrLf
        Else
            .MsgNoteText = .MsgNoteText & "Gender:         " & "Female" & vbCrLf
        End If
        If chkAdopted.Value Then
            .MsgNoteText = .MsgNoteText & "Adopted         " & "Yes" & vbCrLf
        Else
            .MsgNoteText = .MsgNoteText & "Adopted         " & "No" & vbCrLf
        End If
        .MsgNoteText = .MsgNoteText & "DOB:            " & txtDOB & vbCrLf
        .MsgNoteText = .MsgNoteText & "Place of Birth: " & txtPlaceOB & vbCrLf
        .MsgNoteText = .MsgNoteText & "Baptised on:    " & txtBaptised & vbCrLf
        .MsgNoteText = .MsgNoteText & "   In Curch:    " & txtBaptChurch & vbCrLf
        If chkDeceased Then
            .MsgNoteText = .MsgNoteText & "Date Died:      " & txtDateofDeath & vbCrLf
            .MsgNoteText = .MsgNoteText & "Place of Death: " & txtPlaceofDeath & vbCrLf
            .MsgNoteText = .MsgNoteText & "Buried at:      " & txtBuriedAt & vbCrLf
        End If
        sAdd = txtAddress & vbCrLf & txtTown & vbCrLf & txtCounty & vbCrLf & txtPostcode
        .MsgNoteText = .MsgNoteText & "Address:        " & vbCrLf & vbCrLf
        .MsgNoteText = .MsgNoteText & sAdd & vbCrLf & vbCrLf
        .MsgNoteText = .MsgNoteText & "Phone:          " & txtPhone & vbCrLf
        .MsgNoteText = .MsgNoteText & "Email:          " & txtEmail & vbCrLf
        .MsgNoteText = .MsgNoteText & vbCrLf & vbCrLf
        .MsgNoteText = .MsgNoteText & "Spouse(s): " & vbCrLf & vbCrLf
        On Error Resume Next
        For idx = 0 To 4
            .MsgNoteText = .MsgNoteText & lblSpouse(idx) & "    Married on: " & txtMarriageDate(idx) & vbCrLf
        Next idx
        .MsgNoteText = .MsgNoteText & vbCrLf & "Children: " & vbCrLf & vbCrLf
        For idx = 0 To 16
            .MsgNoteText = .MsgNoteText & lblChild(idx) & "    Born on: " & lblChildDOB(idx) & vbCrLf
        Next idx
        On Error GoTo ErrSub
        .MsgNoteText = .MsgNoteText & vbCrLf & "==========================================================" & vbCrLf & vbCrLf
        .MsgNoteText = .MsgNoteText & "Please add further information below here."
        .RecipIndex = 0
        .RecipDisplayName = GetOption(eEmailChanges)
'        .RecipIndex = 1
'        .RecipType = mapBccList
'        .RecipDisplayName = "alternative Email address here"
        On Error GoTo email_box_closed  ' ignore error if user attepmts to close Email box
        .Send True ' popup Email box, comment out True to sent Email directly
    End With
    
Exit Sub
email_box_closed:
   'ignore
   
Exit Sub
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function mnuFileAddChild_Click"
            
    Call Showerror(sErr, 0)

End Sub

Private Sub mnuGotoAdd_Click()
    Select Case ePersonToAdd
        Case eFather
            mnuFileAddFather_Click
        Case eMother
            mnuFileAddMother_Click
        Case eWife, eHusband
            mnuFileAddSpouse_Click
        Case eChild
            mnuFileAddChild_Click
    End Select
End Sub

Private Sub mnuGotoInd_Click()
    If mNewInd <> 0 Then
        Call GetIndividual(mNewInd)
    End If
End Sub

Private Sub mnuGotoRemove_Click()
Dim i As Integer

    If MsgBox("Delete this relationship?", vbYesNo Or vbQuestion, Me.Caption) = vbYes Then
        If Val(lblFather.Tag) = mDelRel Then
            lblFather = ""
            lblFather.Tag = ""
            Call UpdateRelationShips
            Exit Sub
        End If
        If Val(lblMother.Tag) = mDelRel Then
            lblMother = ""
            lblMother.Tag = ""
            Call UpdateRelationShips
            Exit Sub
        End If
        For i = 0 To lblChild.Count - 1
            If Val(lblChild(i).Tag) = mDelRel Then
                lblChild(i).Caption = ""
                lblChildDOB(i).Caption = ""
                Call UpdateRelationShips
                lblChild(i).Tag = ""
                Exit Sub
            End If
        Next i
        For i = 0 To lblSpouse.Count - 1
            If Val(lblSpouse(i).Tag) = mDelRel Then
                lblSpouse(i).Caption = ""
                txtMarriageDate(i).Text = ""
                Call UpdateRelationShips
                lblSpouse(i).Tag = ""
                Exit Sub
            End If
        Next i
    End If
End Sub

Private Sub mnuHelpTopic_Click()
    Call ShowHelpContents(Me.hWnd, HelpConstants.cdlHelpContext, Me.HelpContextID)
End Sub

Private Sub mnuPopProperties_Click()
    If frmPicInfo.invoke(Val(picGallery(mIndex).Tag)) Then
        Call PopulatePicBox(mIndex, picGallery(mIndex).Tag, "", "")
    End If
End Sub

Private Sub mnuPopView_Click()
    Call picGallery_DblClick(mIndex)
End Sub

Private Sub mnuViewGallery_Click()
    eView = eViewType.eGallery
    DoEvents
    picGallery(0).Visible = False
    lblImgCaption(0).Visible = False
    fraOther.Top = 810
    fraOther.Left = 0
    fraOther.ZOrder 0
    lblOtherDetailsName = lblFullName
    Call GetImages
End Sub

Private Sub mnuViewIndividual_Click()
    eView = eViewType.eDetails
    Call GetIndividual(Val(txtSurname.Tag))
    fraDetails.Top = 810
    fraDetails.Left = 0
    fraDetails.ZOrder 0
    fraPedigree.Top = 10000
    fraOther.Top = 10000
End Sub

Private Sub mnuViewPedigree_Click()
    eView = eViewType.ePedigree
    lblPedigreeName = lblFullName
    Call GetPedigree(Val(txtSurname.Tag))
    fraPedigree.Top = 810
    fraPedigree.Left = 0
    fraPedigree.ZOrder 0
    fraDetails.Top = 10000
    fraOther.Top = 10000
End Sub

Private Sub optGender_Click(Index As Integer)
    Call SwitchControls(ONN)
End Sub

Private Sub picGallery_DblClick(Index As Integer)
    If picGallery(Index).Tag <> "" Then
        Call frmPicViewer.invoke(picGallery(Index).Tag, lblImgCaption(Index).Tag, lblImgCaption(Index).Caption, lblOtherDetailsName.Caption)
    End If
End Sub

Private Sub picGallery_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        mIndex = Index
        PopupMenu mnuPop
    End If
End Sub

Private Sub rtbMemo_Change()
    SwitchControls (ONN)
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim lNewId As Long

    On Error Resume Next
    Select Case Button.Key
        Case "Home"
            mnuViewIndividual_Click
            Call GetIndividual(GetOption(eMainIndId))
        Case "Back"
            If mnHistIdx > 0 Then
                bBackorFwd = True
                mnHistIdx = mnHistIdx - 1
                Call GetIndividual(mlHistory(mnHistIdx))
                If eView = eViewType.ePedigree Then
                    GetPedigree (mlHistory(mnHistIdx))
                End If
            End If
        Case "Forward"
            If mnHistIdx < UBound(mlHistory) - 1 Then
                bBackorFwd = True
                mnHistIdx = mnHistIdx + 1
                Call GetIndividual(mlHistory(mnHistIdx))
                If eView = eViewType.ePedigree Then
                    GetPedigree (mlHistory(mnHistIdx))
                End If
            End If
        Case "Index"
            lNewId = frmIndex.invoke(1000, 2099, "")
            If lNewId > 0 Then
                GetIndividual (lNewId)
                Me.HelpContextID = 3
                mnuViewIndividual_Click
            End If
        Case "Details"
            Me.HelpContextID = 3
            mnuViewIndividual_Click
        Case "Pedigree"
            Me.HelpContextID = 4
            mnuViewPedigree_Click
        Case "Pictures"
            Me.HelpContextID = 5
            Call mnuViewGallery_Click
        Case "Email"
            mnuFileSendTo_Click
        Case "Help"
            mnuHelpContents_Click
        Case "Quit"
            mnuFileExit_Click
        Case "New"
            'ToDo: Add 'New' button code.
            If mbChanged Then
                If MsgBox("Ok to save the changes?", vbYesNo Or vbQuestion, Me.Caption) = vbYes Then
                    Call SaveIndividual
                End If
            End If
            ClearIndFrame
            lNewId = frmNewPerson.invoke(0, "", eNone, "", "")
            Call GetIndividual(lNewId)
        
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpContents_Click()
    Call ShowHelpContents(Me.hWnd, HelpConstants.cdlHelpContents, 2)
End Sub

Private Sub mnuHelpIndex_Click()
    Call ShowHelpContents(Me.hWnd, HelpConstants.cdlHelpKey, 0)
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    'ToDo: Add 'mnuWindowArrangeIcons_Click' code.
    MsgBox "Add 'mnuWindowArrangeIcons_Click' code."
End Sub

Private Sub mnuWindowTileVertical_Click()
    'ToDo: Add 'mnuWindowTileVertical_Click' code.
    MsgBox "Add 'mnuWindowTileVertical_Click' code."
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    'ToDo: Add 'mnuWindowTileHorizontal_Click' code.
    MsgBox "Add 'mnuWindowTileHorizontal_Click' code."
End Sub

Private Sub mnuWindowCascade_Click()
    'ToDo: Add 'mnuWindowCascade_Click' code.
    MsgBox "Add 'mnuWindowCascade_Click' code."
End Sub

Private Sub mnuWindowNewWindow_Click()
    'ToDo: Add 'mnuWindowNewWindow_Click' code.
    MsgBox "Add 'mnuWindowNewWindow_Click' code."
End Sub

Private Sub mnuToolsOptions_Click()
    frmOptions.InvokeOptions
End Sub

Private Sub mnuViewWebBrowser_Click()
    'ToDo: Add 'mnuViewWebBrowser_Click' code.
    MsgBox "Add 'mnuViewWebBrowser_Click' code."
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
    'ToDo: Add 'mnuViewRefresh_Click' code.
    MsgBox "Add 'mnuViewRefresh_Click' code."
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbMain.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
End Sub

Private Sub mnuEditPaste_Click()
    'ToDo: Add 'mnuEditPaste_Click' code.
    MsgBox "Add 'mnuEditPaste_Click' code."
End Sub

Private Sub mnuEditCopy_Click()
    'ToDo: Add 'mnuEditCopy_Click' code.
    MsgBox "Add 'mnuEditCopy_Click' code."
End Sub

Private Sub mnuEditCut_Click()
    'ToDo: Add 'mnuEditCut_Click' code.
    MsgBox "Add 'mnuEditCut_Click' code."
End Sub

Private Sub mnuEditUndo_Click()
    'ToDo: Add 'mnuEditUndo_Click' code.
    MsgBox "Add 'mnuEditUndo_Click' code."
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFilePrint_Click()
    'ToDo: Add 'mnuFilePrint_Click' code.
    MsgBox "Add 'mnuFilePrint_Click' code."
End Sub

Private Sub mnuFilePrintPreview_Click()
    'ToDo: Add 'mnuFilePrintPreview_Click' code.
    MsgBox "Add 'mnuFilePrintPreview_Click' code."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub

Private Sub mnuFileSaveAll_Click()
    'ToDo: Add 'mnuFileSaveAll_Click' code.
    MsgBox "Add 'mnuFileSaveAll_Click' code."
End Sub

Private Sub mnuFileSaveAs_Click()
    'ToDo: Add 'mnuFileSaveAs_Click' code.
    MsgBox "Add 'mnuFileSaveAs_Click' code."
End Sub

Private Sub mnuFileSave_Click()
    Call SaveIndividual
End Sub

Private Sub mnuFileClose_Click()
    'ToDo: Add 'mnuFileClose_Click' code.
    MsgBox "Add 'mnuFileClose_Click' code."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'ToDo: add code to process the opened file

End Sub


Private Sub mnuFileNew_Click()
    'ToDo: Add 'mnuFileNew_Click' code.
    MsgBox "Add 'mnuFileNew_Click' code."
End Sub


Private Function GetIndividual(lngIndId As Long) As Boolean
'This function gets details about an individual and
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sErr As String
Dim idx As Long
Dim iResp As Integer

    On Error GoTo ErrSub
    
    If mbChanged Then
        If MsgBox("Ok to save the changes?", vbYesNo Or vbQuestion, Me.Caption) = vbYes Then
            Call SaveIndividual
        End If
    End If

    Call ClearIndFrame

    SQL = "SELECT * FROM " & gtcINDIVIDUALS & " WHERE " & _
            gccINDID & " = " & lngIndId
            
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    If Not RS.EOF And Not RS.BOF Then
        txtSurname = Format(RS(gccINDSURNAME))
        txtSurname.Tag = RS(gccINDID)
        txtFirstNames = Format(RS(gccINDFIRSTNAMES))
        lblFullName = Trim(Trim(RS(gccINDFIRSTNAMES)) & " " & Trim(RS(gccINDSURNAME)))
        If Not IsNull(RS(gccINDDOBTEXT)) Then
            txtDOB = RS(gccINDDOBTEXT)
        Else
            txtDOB = ""
        End If
        If RS(gccINDADOPTED) Then
            chkAdopted.Value = vbChecked
        Else
            chkAdopted.Value = vbUnchecked
        End If
        'chkAdopted.Value = CInt(RS(gccINDADOPTED))
        txtPlaceOB = Format(RS(gccINDPLACEOFBIRTH))
        If Not IsNull(RS(gccINDBAPTDATETEXT)) Then
            txtBaptised = RS(gccINDBAPTDATETEXT)
        Else
            txtBaptChurch = ""
        End If
        txtBaptChurch = Format(RS(gccINDBAPTCHURCH))
        chkDeceased = Abs(RS(gccINDDECEASED))
        If RS(gccINDGENDER) = "F" Then
            optGender(1).Value = True
        Else
            optGender(0).Value = True
        End If
        If RS(gccINDDECEASED) Then
            txtDateofDeath.Enabled = True
            txtPlaceofDeath.Enabled = True
            txtBuriedAt.Enabled = True
            txtDateofDeath = RS(gccINDDEATHDATETEXT)
            txtPlaceofDeath = Format(RS(gccINDPLACEOFDEATH))
            txtBuriedAt = Format(RS(gccINDPLACEBURIED))
            If RS(gccINDDEATHDATEDATE) = 0 Then
                lblAge = "(Unk)"
            Else
                lblAge = GetAge(RS(gccINDDOBDATE), RS(gccINDDEATHDATEDATE))
            End If
        Else
            lblAge = GetAge(RS(gccINDDOBDATE), RS(gccINDDEATHDATEDATE))
            txtDateofDeath.Enabled = False
            txtPlaceofDeath.Enabled = False
            txtBuriedAt.Enabled = False
        End If
        txtAddress = RS(gccINDADDLINE1) & vbCrLf & RS(gccINDADDLINE2) & vbCrLf & RS(gccINDADDLINE3)
        txtTown = Format(RS(gccINDTOWN))
        txtCounty = Format(RS(gccINDCOUNTY))
        txtPostcode = Format(RS(gccINDPOSTCODE))
        txtPhone = Format(RS(gccINDPHONE))
        txtEmail = Format(RS(gccINDEMAIL))
        If RS(gccINDFATHERID) <> 0 Then
            lblFather.Tag = RS(gccINDFATHERID)
            lblFather = GetFullName(RS(gccINDFATHERID))
        Else
            lblFather = ""
            lblFather.Tag = ""
        End If
        If RS(gccINDMOTHERID) <> 0 Then
            lblMother.Tag = RS(gccINDMOTHERID)
            lblMother = GetFullName(RS(gccINDMOTHERID))
        Else
            lblMother = ""
            lblMother.Tag = ""
        End If
        
        rtbMemo.TextRTF = ""
        
        On Error Resume Next
        rtbMemo.TextRTF = RS(gccINDMEMO)
        On Error GoTo ErrSub
        
        Call GetSpouses(RS(gccINDID))
        Call GetChildren(RS(gccINDID))
        
        For idx = 0 To 6
            If RS(gccINDDOBDATE) < (1842 + (idx * 10)) * 10000 Then
                cmdCensus(idx).Enabled = True
                If CensusExists(RS(gccINDID), 1841 + (idx * 10)) Then
                    cmdCensus(idx).FontBold = True
                Else
                    cmdCensus(idx).FontBold = False
                End If
            Else
                cmdCensus(idx).Enabled = False
            End If
        Next idx
        
        If PicturesExist(RS(gccINDID)) Then
            cmdOther.FontBold = True
        Else
            cmdOther.FontBold = False
        End If
        
        GetIndividual = True
        
    End If
    
    If Not bBackorFwd Then
        If UBound(mlHistory) >= 1 Then
            If Val(txtSurname.Tag) <> mlHistory(UBound(mlHistory) - 1) Then
                ReDim Preserve mlHistory(UBound(mlHistory) + 1)
                mlHistory(UBound(mlHistory) - 1) = Val(txtSurname.Tag)
                mnHistIdx = UBound(mlHistory) - 1
            End If
        Else
            mlHistory(0) = Val(txtSurname.Tag)
        End If
    End If
    
    bBackorFwd = False
    
    Call SwitchControls(OFF)

Exit Function
ErrSub:

    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function GetIndividual"
            
    Call Showerror(sErr, 0)


End Function


Private Function CensusExists(lngId As Long, iYear As Integer) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sErr As String


    On Error GoTo ErrSub
    
    SQL = "SELECT count(*) as numrecs FROM " & gtcCENSUS & _
            " LEFT JOIN " & gtcCENSUSHEADER & " ON " & _
            gtcCENSUS & "." & gccCENCNHID & " = " & gtcCENSUSHEADER & "." & gccCNHID & _
            " WHERE " & _
            gccCENINDID & " = " & lngId & " AND " & _
            gccCNHYEAR & " = " & iYear
        
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    If RS("numrecs") <> 0 Then
        CensusExists = True
    End If

Exit Function
ErrSub:

    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function CensusExists"
            
    Call Showerror(sErr, 0)


End Function

Private Function PicturesExist(lngId As Long) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sErr As String


    On Error GoTo ErrSub
    
    SQL = "SELECT count(*) as numrecs FROM Imagelink Where ImlIndID = " & lngId
        
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    If RS("numrecs") <> 0 Then
        PicturesExist = True
    End If

Exit Function
ErrSub:

    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function PicturesExist"
            
    Call Showerror(sErr, 0)


End Function


Private Function GetChildren(lngId As Long) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim i As Integer
Dim sGender As String
Dim sErr As String

    On Error GoTo ErrSub

    sGender = IndGender(lngId)


    If sGender = "M" Then
        SQL = "SELECT * FROM " & gtcINDIVIDUALS & _
                " WHERE " & gccINDFATHERID & " = " & lngId & _
                " ORDER BY " & gccINDDOBDATE
    Else
        SQL = "SELECT * FROM " & gtcINDIVIDUALS & _
                " WHERE " & gccINDMOTHERID & " = " & lngId & _
                " ORDER BY " & gccINDDOBDATE
    End If
            
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    i = 0
    lvChildren.ColumnHeaders.Clear
    lvChildren.ListItems.Clear

    Call lvChildren.ColumnHeaders.Add(1, "Name", "Children", lvChildren.Width - 60)
'    lvSpouses.ColumnHeaders(1).Tag = lvwAscending
'    Call lvSpouses.ColumnHeaders.Add(2, "Dob.", "Description", 700)
    lvChildren.View = lvwReport
    
    If eView = eViewType.ePedigree Then
        i = 1
    Else
        i = 0
    End If
    
    Do While Not RS.EOF
        If eView = eViewType.ePedigree Then
            Call lvChildren.ListItems.Add(i, RS(gccINDID) & "X", Trim(Trim(RS(gccINDFIRSTNAMES)) & " " & Trim(RS(gccINDSURNAME))))
'            Itmx.SubItems(1) = RS(gccINDDOBTEXT)
        Else
            If i > 0 Then
                Load lblChild(i)
                lblChild(i).Top = lblChild(i - 1).Top + lblChild(i - 1).Height
                Load lblChildDOB(i)
                lblChildDOB(i).Top = lblChild(i).Top
            End If
            lblChild(i).Visible = True
            lblChild(i) = Trim(Trim(RS(gccINDFIRSTNAMES)) & " " & Trim(RS(gccINDSURNAME)))
            lblChildDOB(i).Visible = True
            lblChildDOB(i) = "b. " & RS(gccINDDOBTEXT)
            lblChild(i).Tag = RS(gccINDID)
            If RS(gccINDGENDER) = "F" Then
                lblChild(i).BackColor = &HE7DEFE
            Else
                lblChild(i).BackColor = &HFEFFD7
            End If
        End If
        RS.MoveNext
        i = i + 1
    Loop
    
    RS.Close
    
Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function GetChildren"
            
    Call Showerror(sErr, 0)

End Function

Private Function GetSpouses(lngId As Long) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim i As Integer
Dim Itmx As ListItem
Dim sGender As String
Dim sErr As String

    On Error GoTo ErrSub

    sGender = IndGender(lngId)

    If sGender = "M" Then
        SQL = "SELECT * FROM " & gtcMARRIAGES & _
                " WHERE " & gccSPOHUSBANDID & " = " & lngId & _
                " ORDER BY " & gccSPOMARRIAGEDATEDATE & " asc, " & gccSPOWIFEID & " asc"
    Else
        SQL = "SELECT * FROM " & gtcMARRIAGES & _
                " WHERE " & gccSPOWIFEID & " = " & lngId & _
                " ORDER BY " & gccSPOMARRIAGEDATEDATE & " asc, " & gccSPOHUSBANDID & " asc"
    End If
            
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    lvSpouses.ColumnHeaders.Clear
    lvSpouses.ListItems.Clear

    Call lvSpouses.ColumnHeaders.Add(1, "Name", "Spouses", lvSpouses.Width - 60)
'    lvSpouses.ColumnHeaders(1).Tag = lvwAscending
'    Call lvSpouses.ColumnHeaders.Add(2, "Dob.", "Description", 700)
    lvSpouses.View = lvwReport
    
    If eView = eViewType.ePedigree Then
        i = 1
    Else
        i = 0
    End If
    
    Do While Not RS.EOF
        If eView = eViewType.ePedigree Then
            If sGender = "M" Then
                Call lvSpouses.ListItems.Add(i, RS(gccSPOWIFEID) & "X", GetFullName(RS(gccSPOWIFEID)))
            Else
                Call lvSpouses.ListItems.Add(i, RS(gccSPOHUSBANDID) & "X", GetFullName(RS(gccSPOHUSBANDID)))
            End If
'            Itmx.SubItems(1) = RS(gccINDDOBTEXT)
        Else
            If i > 0 Then
                Load lblSpouse(i)
                lblSpouse(i).Top = lblSpouse(i - 1).Top + lblSpouse(i - 1).Height
                Load lblSpouseNum(i)
                lblSpouseNum(i).Top = lblSpouse(i).Top
                Load txtMarriageDate(i)
                txtMarriageDate(i).Top = lblSpouse(i).Top
                lblSpouseNum(i).Caption = i + 1
            End If
            lblSpouseNum(i).Visible = True
            lblSpouse(i).Visible = True
            txtMarriageDate(i).Visible = True
'            cmdViewMarriageCert(i).Visible = True
            If optGender(0).Value = True Then
                lblSpouse(i) = GetFullName(RS(gccSPOWIFEID))
                lblSpouse(i).Tag = RS(gccSPOWIFEID)
            Else
                lblSpouse(i) = GetFullName(RS(gccSPOHUSBANDID))
                lblSpouse(i).Tag = RS(gccSPOHUSBANDID)
            End If
            txtMarriageDate(i) = RS(gccSPOMARRIAGEDATETEXT)
            If sGender = "F" Then
                lblSpouse(i).BackColor = &HFEFFD7
            Else
                lblSpouse(i).BackColor = &HE7DEFE
            End If
            
        End If
        i = i + 1
        RS.MoveNext
    Loop
    
    RS.Close
    
Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function GetSpouses"
            
    Call Showerror(sErr, 0)


End Function

Private Sub ClearIndFrame()
Dim i As Integer

    lblChild(0).Caption = ""
    lblChild(0).Tag = ""
    lblChildDOB(0).Caption = ""

    For i = lblChild.Count - 1 To 1 Step -1
        Unload lblChild(i)
        Unload lblChildDOB(i)
    Next i
    
    lblSpouse(0) = ""
    lblSpouse(0).Tag = ""
    txtMarriageDate(0).Text = ""
    
    For i = lblSpouse.Count - 1 To 1 Step -1
        Unload lblSpouse(i)
        Unload lblSpouseNum(i)
        Unload txtMarriageDate(i)
    Next i
        
    lblFullName = ""
    lblFather = ""
    lblMother = ""
    txtSurname = ""
    txtSurname.Tag = ""
    txtFirstNames = ""
    optGender(0).Value = True
    chkAdopted.Value = vbUnchecked
    txtDOB = ""
    txtPlaceOB = ""
    txtBaptised = ""
    txtBaptChurch = ""
    chkDeceased.Value = vbUnchecked
    txtDateofDeath = ""
    txtPlaceofDeath = ""
    txtBuriedAt = ""
    txtAddress = ""
    txtTown = ""
    txtCounty = ""
    txtPostcode = ""
    txtPhone = ""
    txtEmail = ""
    mbChanged = False
    
End Sub

Private Function SaveIndividual() As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sAdd() As String

Dim sErr As String

    On Error GoTo ErrSub

    If ValidDetails Then
        If Val(txtSurname.Tag) <> 0 Then
            SQL = "SELECT * FROM " & gtcINDIVIDUALS & " WHERE " & _
                    gccINDID & " = " & Val(txtSurname.Tag)
        Else
            'Create a blank recordset
            SQL = "SELECT * FROM " & gtcINDIVIDUALS & " WHERE 1 = 2"
        End If
        
        Set RS = New ADODB.Recordset
        
        RS.Open SQL, gApp.cn, adOpenKeyset, adLockPessimistic
        
        If Val(txtSurname.Tag) = 0 Then
            RS.AddNew
        Else
            If RS.RecordCount < 1 Then
                MsgBox "ERROR This record has not been found on the database!", vbOKOnly Or vbCritical, Me.Caption
                RS.Close
                Exit Function
            End If
        End If
        RS(gccINDSURNAME) = Trim(txtSurname)
        RS(gccINDFIRSTNAMES) = Trim(txtFirstNames)
        If optGender(1).Value = True Then
            RS(gccINDGENDER) = "F"
        Else
            RS(gccINDGENDER) = "M"
        End If
        RS(gccINDDOBTEXT) = Trim(txtDOB)
        RS(gccINDDOBDATE) = Val(txtDOB.Tag)
        RS(gccINDPLACEOFBIRTH) = Trim(txtPlaceOB)
        RS(gccINDADOPTED) = chkAdopted.Value
        RS(gccINDBAPTDATETEXT) = Trim(txtBaptised)
        RS(gccINDBAPTDATEDATE) = Val(txtBaptised.Tag)
        RS(gccINDBAPTCHURCH) = Trim(txtBaptChurch)
        RS(gccINDDECEASED) = CBool(chkDeceased.Value)
        RS(gccINDDEATHDATETEXT) = Trim(txtDateofDeath)
        RS(gccINDDEATHDATEDATE) = Val(txtDateofDeath.Tag)
        RS(gccINDPLACEOFDEATH) = Trim(txtPlaceofDeath)
        RS(gccINDPLACEBURIED) = Trim(txtBuriedAt)
        sAdd = Split(Trim(txtAddress), vbCrLf)
        If UBound(sAdd) >= 0 Then RS(gccINDADDLINE1) = sAdd(0)
        If UBound(sAdd) >= 1 Then RS(gccINDADDLINE2) = sAdd(1)
        If UBound(sAdd) >= 2 Then RS(gccINDADDLINE3) = sAdd(2)
        RS(gccINDTOWN) = Trim(txtTown)
        RS(gccINDCOUNTY) = Trim(txtCounty)
        RS(gccINDPOSTCODE) = Trim(txtPostcode)
        RS(gccINDPHONE) = Trim(txtPhone)
        RS(gccINDEMAIL) = Trim(txtEmail)
        RS(gccINDFATHERID) = Val(lblFather.Tag)
        RS(gccINDMOTHERID) = Val(lblMother.Tag)
        RS(gccINDMEMO) = rtbMemo.TextRTF
        RS.Update
        RS.Close
        SwitchControls (OFF)
        SaveIndividual = True
        UpdateRelationShips
    End If

Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function SaveIndividual"
            
    Call Showerror(sErr, 0)

End Function

Private Function ValidDetails() As Boolean
Dim sMess As String
Dim sDate As String
Dim lDateNum As Long
Dim i As Integer

    If Trim(txtSurname) = "" Then
        sMess = sMess & "You must enter a surname." & vbCrLf
    End If
    
    sDate = Trim(txtDOB)
    If sDate = "" Then
        sMess = sMess & "You must indicate an approx date of birth." & vbCrLf
    Else
        lDateNum = ValidDate(sDate)
        If lDateNum = 0 Then
            sMess = sMess & "The date of birth is not recognised as a valid date format." & vbCrLf
        Else
            txtDOB = sDate
            txtDOB.Tag = lDateNum
        End If
    End If
    
    sDate = Trim(txtBaptised)
    If sDate = "" Then
        txtBaptised.Tag = 0
    Else
        lDateNum = ValidDate(sDate)
        If lDateNum = 0 Then
            sMess = sMess & "The date of baptism is not recognised as a valid date format." & vbCrLf
        Else
            txtBaptised = sDate
            txtBaptised.Tag = lDateNum
        End If
    End If
    
    sDate = Trim(txtDateofDeath)
    If sDate = "" Then
        txtDateofDeath.Tag = 0
    Else
        lDateNum = ValidDate(sDate)
        If lDateNum = 0 Then
            sMess = sMess & "The date of death is not recognised as a valid date format." & vbCrLf
        Else
            txtDateofDeath = sDate
            txtDateofDeath.Tag = lDateNum
        End If
    End If

    For i = 0 To lblSpouse.Count - 1
        sDate = Trim(txtMarriageDate(i))
        If sDate = "" Then
            txtMarriageDate(i).Tag = 0
        Else
            lDateNum = ValidDate(sDate)
            If lDateNum = 0 Then
                sMess = sMess & "The date of marriage is not recognised as a valid date format." & vbCrLf
            Else
                txtMarriageDate(i) = sDate
                txtMarriageDate(i).Tag = lDateNum
            End If
        End If
    Next i

    If sMess = "" Then
        ValidDetails = True
    Else
        MsgBox "You cannot save this data because of the following errors..." & vbCrLf & vbCrLf & sMess, vbOKOnly Or vbCritical, Me.Caption
    End If

End Function

Private Sub tbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Father"
            mnuFileAddFather_Click
        Case "Mother"
            mnuFileAddMother_Click
        Case "Spouse"
            mnuFileAddSpouse_Click
        Case "Child"
            mnuFileAddChild_Click
        Case "Unrelated"
            mnuFileAddUnrelated_Click
    End Select
End Sub

Private Sub txtAddress_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub txtBaptChurch_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub txtBaptised_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub txtBuriedAt_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub lblchild_Click(Index As Integer)
    If Val(lblChild(Index).Tag) <> 0 Then
        Call GetIndividual(Val(lblChild(Index).Tag))
    End If
End Sub

Private Sub lblchild_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    HighLightChild (Index)
End Sub

Private Sub txtCounty_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub txtDateofDeath_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub txtDOB_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub txtEmail_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub txtFFather_Click()
    Call GetPedigree(Val(txtFFather.Tag))
End Sub

Private Sub txtFFFather_Click()
    Call GetPedigree(Val(txtFFFather.Tag))
End Sub

Private Sub txtFFMother_Click()
    Call GetPedigree(Val(txtFFMother.Tag))
End Sub

Private Sub txtFirstNames_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub txtFMFather_Click()
    Call GetPedigree(Val(txtFMFather.Tag))
End Sub

Private Sub txtFMMother_Click()
    Call GetPedigree(Val(txtFMMother.Tag))
End Sub

Private Sub txtFMother_Click()
    Call GetPedigree(Val(txtFMother.Tag))
End Sub

Private Sub txtMainInd_GotFocus()
    On Error Resume Next
    txtInvisible.SetFocus
End Sub

Private Sub txtMarriageDate_Change(Index As Integer)
    Call SwitchControls(ONN)
End Sub

Private Sub txtMFather_Click()
    Call GetPedigree(Val(txtMFather.Tag))
End Sub

Private Sub txtMFFather_Click()
    Call GetPedigree(Val(txtMFFather.Tag))
End Sub

Private Sub txtMFMother_Click()
    Call GetPedigree(Val(txtMFMother.Tag))
End Sub

Private Sub txtMMFather_Click()
    Call GetPedigree(Val(txtMMFather.Tag))
End Sub

Private Sub txtMMMother_Click()
    Call GetPedigree(Val(txtMMMother.Tag))
End Sub

Private Sub txtMMother_Click()
    Call GetPedigree(Val(txtMMother.Tag))
End Sub

Private Sub txtPedFather_Click()
    Call GetPedigree(Val(txtPedFather.Tag))
End Sub

Private Sub txtPedMother_Click()
    Call GetPedigree(Val(txtPedMother.Tag))
End Sub

Private Sub txtPhone_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub txtPlaceOB_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub txtPlaceofDeath_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub txtPostcode_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub txtSurname_Change()
    Call SwitchControls(ONN)
End Sub

Private Sub SwitchControls(bState As Boolean)
Dim i As Integer

    mbChanged = bState
    
'    For i = 0 To lblSpouse.Count - 1
'        cmdViewMarriageCert(i).Enabled = Not bState
'    Next i
    
End Sub

Private Sub txtTown_Change()
    Call SwitchControls(ONN)
End Sub

Private Function GetPedigree(lMainId As Long) As Boolean
'This function populates the pedigree chart for the Id of the person chosen


    If lMainId = 0 Then Exit Function
    
    'Always keep the Details page in sync
    GetIndividual (lMainId)
    
    txtMainInd = "": txtMainInd.Tag = ""
    txtPedFather = "": txtPedFather.Tag = ""
    txtPedMother = "": txtPedMother.Tag = ""
    
    txtFFather = "": txtFFather.Tag = ""
    txtFMother = "": txtFMother.Tag = ""
    txtMFather = "": txtMFather.Tag = ""
    txtMMother = "": txtMMother.Tag = ""

    txtFFFather = "": txtFFFather.Tag = ""
    txtFFMother = "": txtFFMother.Tag = ""
    txtFMFather = "": txtFMFather.Tag = ""
    txtFMMother = "": txtFMMother.Tag = ""
    
    txtMFFather = "": txtMFFather.Tag = ""
    txtMFMother = "": txtMFMother.Tag = ""
    txtMMFather = "": txtMMFather.Tag = ""
    txtMMMother = "": txtMMMother.Tag = ""

    txtMainInd.Tag = lMainId
    txtMainInd = GetFullName(lMainId, True, True)
    
    lblPedigreeName = lblFullName
    
'Parents
    txtPedFather.Tag = GetParent(lMainId, eFather)
    If Val(txtPedFather.Tag) <> 0 Then txtPedFather = GetFullName(Val(txtPedFather.Tag), True, True)
    
    txtPedMother.Tag = GetParent(lMainId, eMother)
    If Val(txtPedMother.Tag) <> 0 Then txtPedMother = GetFullName(Val(txtPedMother.Tag), True, True)

'Paternal Grandparents
    txtFFather.Tag = GetParent(Val(txtPedFather.Tag), eFather)
    If Val(txtFFather.Tag) <> 0 Then txtFFather = GetFullName(Val(txtFFather.Tag), True, True)
    
    txtFMother.Tag = GetParent(Val(txtPedFather.Tag), eMother)
    If Val(txtFMother.Tag) <> 0 Then txtFMother = GetFullName(Val(txtFMother.Tag), True, True)

'Maternal Grandparents
    txtMFather.Tag = GetParent(Val(txtPedMother.Tag), eFather)
    If Val(txtMFather.Tag) <> 0 Then txtMFather = GetFullName(Val(txtMFather.Tag), True, True)
    
    txtMMother.Tag = GetParent(Val(txtPedMother.Tag), eMother)
    If Val(txtMMother.Tag) <> 0 Then txtMMother = GetFullName(Val(txtMMother.Tag), True, True)

'Fathers Paternal Grandparents
    txtFFFather.Tag = GetParent(Val(txtFFather.Tag), eFather)
    If Val(txtFFFather.Tag) <> 0 Then txtFFFather = GetFullName(Val(txtFFFather.Tag), True, True)
    
    txtFFMother.Tag = GetParent(Val(txtFFather.Tag), eMother)
    If Val(txtFFMother.Tag) <> 0 Then txtFFMother = GetFullName(Val(txtFFMother.Tag), True, True)

'Fathers Maternal Grandparents
    txtFMFather.Tag = GetParent(Val(txtFMother.Tag), eFather)
    If Val(txtFMFather.Tag) <> 0 Then txtFMFather = GetFullName(Val(txtFMFather.Tag), True, True)
    
    txtFMMother.Tag = GetParent(Val(txtFMother.Tag), eMother)
    If Val(txtFMMother.Tag) <> 0 Then txtFMMother = GetFullName(Val(txtFMMother.Tag), True, True)

'Mothers Paternal Grandparents
    txtMFFather.Tag = GetParent(Val(txtMFather.Tag), eFather)
    If Val(txtMFFather.Tag) <> 0 Then txtMFFather = GetFullName(Val(txtMFFather.Tag), True, True)
    
    txtMFMother.Tag = GetParent(Val(txtMFather.Tag), eMother)
    If Val(txtMFMother.Tag) <> 0 Then txtMFMother = GetFullName(Val(txtMFMother.Tag), True, True)

'Mothers Maternal Grandparents
    txtMMFather.Tag = GetParent(Val(txtMMother.Tag), eFather)
    If Val(txtMMFather.Tag) <> 0 Then txtMMFather = GetFullName(Val(txtMMFather.Tag), True, True)
    
    txtMMMother.Tag = GetParent(Val(txtMMother.Tag), eMother)
    If Val(txtMMMother.Tag) <> 0 Then txtMMMother = GetFullName(Val(txtMMMother.Tag), True, True)

    Call GetSpouses(lMainId)
    Call GetChildren(lMainId)
    On Error Resume Next
    txtInvisible.SetFocus
    
End Function

Private Function GetParent(lId As Long, eRel As eRelationships) As Long
Dim RS As ADODB.Recordset
Dim SQL As String

    Select Case eRel
        Case eFather
            SQL = "SELECT " & gccINDFATHERID
        Case eMother
            SQL = "SELECT " & gccINDMOTHERID
    End Select
    SQL = SQL & " FROM " & gtcINDIVIDUALS & " WHERE " & _
            gccINDID & " = " & lId

    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    If Not RS.BOF And Not RS.EOF Then
        If eRel = eFather Then
            GetParent = Val(Format(RS(gccINDFATHERID)))
        Else
            GetParent = Val(Format(RS(gccINDMOTHERID)))
        End If
    End If
    
End Function

Private Function GetImages()
Dim RS As ADODB.Recordset
Dim SQL As String
Dim i As Integer
Dim f As FileAttribute
Dim lPicTop As Long
Dim lLabelTop As Long

    For i = picGallery.Count - 1 To 1 Step -1
        Unload picGallery(i)
        Unload lblImgCaption(i)
    Next i
    DoEvents
            
    picGallery(0).Picture = LoadPicture("")
    picGallery(0).Tag = ""
    picGallery(0).Visible = True
    lblImgCaption(0).Caption = ""
    lblImgCaption(0).Tag = ""
    lblImgCaption(0).Visible = True

    SQL = "SELECT * FROM " & gtcIMAGELINK & _
            " LEFT JOIN " & gtcIMAGES & " ON " & _
            gtcIMAGES & "." & gccIMGID & " = " & gtcIMAGELINK & "." & gccIMLIMGID & _
            " Where " & gccIMLINDID & " = " & Val(txtSurname.Tag) & _
            " Order by " & gccIMGDATEDATE

    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    i = 0
    Do While Not RS.EOF
        If Not IsNull(RS(gccIMGID)) Then
            If i > 0 Then
                Call LoadNewPicBox(i)
            End If
            Call PopulatePicBox(i, RS(gccIMGID), RS(gccIMGNAME), RS(gccIMGCAPTION))
            i = i + 1
            If i > 15 Then Exit Do
        End If
        RS.MoveNext
    Loop
    RS.Close


End Function

Private Function LoadNewPicBox(idx As Integer)

    Load picGallery(idx)
    Load lblImgCaption(idx)
    If idx = 8 Then
        picGallery(idx).Left = 60
        picGallery(idx).Top = 3030
        lblImgCaption(idx).Left = 60
        lblImgCaption(idx).Top = 4410
    Else
        picGallery(idx).Left = picGallery(idx - 1).Left + picGallery(idx).Width + 45
        picGallery(idx).Top = picGallery(idx - 1).Top
        lblImgCaption(idx).Left = picGallery(idx).Left
        lblImgCaption(idx).Top = lblImgCaption(idx - 1).Top
    End If
    picGallery(idx).Visible = True
    lblImgCaption(idx).Visible = True

End Function

Private Function PopulatePicBox(idx As Integer, lImgId As Long, sFileName As String, sCaption As String)
Dim X, Y As Single
Dim lWidth, lHeight As Long
Dim objPic As IPictureDisp
Dim Factor As Single
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sErr As String

    On Error GoTo ErrSub
    
    If sFileName = "" Then
        SQL = "SELECT * FROM " & gtcIMAGES & " WHERE " & gccIMGID & " = " & lImgId

        Set RS = New ADODB.Recordset
        
        RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
        
        If RS.BOF And RS.EOF Then
            Exit Function
        Else
            sFileName = RS(gccIMGNAME)
            sCaption = RS(gccIMGCAPTION)
        End If
        RS.Close
    End If

    lblImgCaption(idx).Caption = Format(sCaption)
    lblImgCaption(idx).Tag = App.Path & "\" & sFileName
    picGallery(idx).Tag = lImgId
    Set objPic = LoadPicture(lblImgCaption(idx).Tag)
    lWidth = Int(objPic.Width)
    lHeight = Int(objPic.Height)
    
    If lHeight > lWidth Then
        Factor = picGallery(idx).ScaleHeight / lHeight
        lHeight = picGallery(idx).ScaleHeight
        lWidth = lWidth * Factor
        X = Int((picGallery(idx).ScaleWidth - lWidth) / 2)
    Else
        Factor = picGallery(idx).ScaleWidth / lWidth
        lWidth = picGallery(idx).ScaleWidth
        lHeight = lHeight * Factor
        Y = Int((picGallery(idx).ScaleHeight - lHeight) / 2)
    End If
    
    picGallery(idx).Picture = LoadPicture()
    picGallery(idx).PaintPicture objPic, X, Y, lWidth, lHeight
    DoEvents

Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function PopulatePicBox"
            
    Call Showerror(sErr, 0)

End Function

Private Sub HighLightChild(idx As Integer)
Dim i As Integer
    
    For i = 0 To lblChild.Count - 1
        If i = idx Then
            lblChild(i).ForeColor = vbBlue
            lblChild(i).FontUnderline = True
        Else
            lblChild(i).ForeColor = vbBlack
            lblChild(i).FontUnderline = False
        End If
    Next i
End Sub

Private Sub HighLightSpouse(idx As Integer)
Dim i As Integer
    
    For i = 0 To lblSpouse.Count - 1
        If i = idx Then
            lblSpouse(i).ForeColor = vbBlue
            lblSpouse(i).FontUnderline = True
        Else
            lblSpouse(i).ForeColor = vbBlack
            lblSpouse(i).FontUnderline = False
        End If
    Next i
End Sub



