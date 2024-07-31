VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F34C6509-63A5-11D3-B1E9-0008C7636E27}#16.0#0"; "UPDOWN.OCX"
Begin VB.Form frmCCSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CC Car Settings"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmCCSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab tabMain 
      Height          =   4125
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7276
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "CC Car Setup"
      TabPicture(0)   =   "frmCCSetup.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraWing"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "CC Car Pitstop Stratergy"
      TabPicture(1)   =   "frmCCSetup.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "updLaps"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin UpDonwOCX.UpDown updLaps 
         Height          =   255
         Left            =   -74400
         TabIndex        =   54
         Top             =   450
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   450
         Text            =   "0"
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
      Begin VB.Frame Frame7 
         Caption         =   "Fuel Load"
         Height          =   735
         Left            =   120
         TabIndex        =   40
         Top             =   3120
         Width           =   2175
         Begin UpDonwOCX.UpDown updFuel 
            Height          =   255
            Left            =   1440
            TabIndex        =   62
            Top             =   270
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            Max             =   32767
            Min             =   0
            MaxLength       =   0
            Text            =   "0"
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
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Fuel Load"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   300
            Width           =   705
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Three Stops"
         Height          =   1370
         Left            =   -74880
         TabIndex        =   24
         Top             =   2640
         Width           =   5655
         Begin UpDonwOCX.UpDown upd3Stops 
            Height          =   255
            Left            =   1320
            TabIndex        =   48
            Top             =   255
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updRange32 
            Height          =   255
            Left            =   4800
            TabIndex        =   47
            Top             =   615
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updRange33 
            Height          =   255
            Left            =   4800
            TabIndex        =   46
            Top             =   975
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updRange31 
            Height          =   255
            Left            =   4800
            TabIndex        =   45
            Top             =   255
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updStart33 
            Height          =   255
            Left            =   3360
            TabIndex        =   44
            Top             =   975
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updStart32 
            Height          =   255
            Left            =   3360
            TabIndex        =   43
            Top             =   615
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updStart31 
            Height          =   255
            Left            =   3360
            TabIndex        =   42
            Top             =   255
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "% doing 3 stops"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   39
            Top             =   285
            Width           =   1110
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Pit within"
            Height          =   195
            Index           =   2
            Left            =   4100
            TabIndex        =   30
            Top             =   1005
            Width           =   630
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Pit within"
            Height          =   195
            Index           =   1
            Left            =   4100
            TabIndex        =   29
            Top             =   645
            Width           =   630
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Pit within"
            Height          =   195
            Index           =   0
            Left            =   4100
            TabIndex        =   28
            Top             =   285
            Width           =   630
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Start pitting at lap"
            Height          =   195
            Index           =   2
            Left            =   2040
            TabIndex        =   27
            Top             =   1005
            Width           =   1230
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Start pitting at lap"
            Height          =   195
            Index           =   1
            Left            =   2040
            TabIndex        =   26
            Top             =   645
            Width           =   1230
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Start pitting at lap"
            Height          =   195
            Index           =   0
            Left            =   2040
            TabIndex        =   25
            Top             =   285
            Width           =   1230
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Two Stops"
         Height          =   1000
         Left            =   -74880
         TabIndex        =   23
         Top             =   1590
         Width           =   5655
         Begin UpDonwOCX.UpDown updRange22 
            Height          =   255
            Left            =   4820
            TabIndex        =   51
            Top             =   615
            Width           =   695
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updStart22 
            Height          =   255
            Left            =   3360
            TabIndex        =   50
            Top             =   615
            Width           =   695
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown upd2Stops 
            Height          =   255
            Left            =   1320
            TabIndex        =   49
            Top             =   255
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updStart21 
            Height          =   255
            Left            =   3360
            TabIndex        =   52
            Top             =   255
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updRange21 
            Height          =   255
            Left            =   4815
            TabIndex        =   53
            Top             =   255
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "% doing 2 stops"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   285
            Width           =   1110
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Start pitting at lap"
            Height          =   195
            Index           =   5
            Left            =   2040
            TabIndex        =   36
            Top             =   645
            Width           =   1230
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Start pitting at lap"
            Height          =   195
            Index           =   4
            Left            =   2040
            TabIndex        =   35
            Top             =   285
            Width           =   1230
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Pit within"
            Height          =   195
            Index           =   5
            Left            =   4100
            TabIndex        =   33
            Top             =   285
            Width           =   630
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Pit within"
            Height          =   195
            Index           =   4
            Left            =   4100
            TabIndex        =   32
            Top             =   645
            Width           =   630
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "One Stop"
         Height          =   700
         Left            =   -74880
         TabIndex        =   22
         Top             =   840
         Width           =   5655
         Begin UpDonwOCX.UpDown upd1Stops 
            Height          =   255
            Left            =   1320
            TabIndex        =   55
            Top             =   255
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updStart11 
            Height          =   255
            Left            =   3360
            TabIndex        =   56
            Top             =   255
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updRange11 
            Height          =   255
            Left            =   4815
            TabIndex        =   57
            Top             =   255
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            Text            =   "0"
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
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "% doing 1 stop"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   285
            Width           =   1035
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Start pitting at lap"
            Height          =   195
            Index           =   3
            Left            =   2040
            TabIndex        =   34
            Top             =   285
            Width           =   1230
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Pit within"
            Height          =   195
            Index           =   3
            Left            =   4100
            TabIndex        =   31
            Top             =   285
            Width           =   630
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Psysics"
         Height          =   1700
         Left            =   120
         TabIndex        =   16
         Top             =   1395
         Width           =   2175
         Begin UpDonwOCX.UpDown updAcc 
            Height          =   285
            Left            =   1440
            TabIndex        =   61
            Top             =   225
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Max             =   100
            Min             =   0
            MaxLength       =   0
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updGrip 
            Height          =   285
            Left            =   1440
            TabIndex        =   60
            Top             =   585
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Max             =   100
            Min             =   0
            MaxLength       =   0
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updAir 
            Height          =   285
            Left            =   1440
            TabIndex        =   59
            Top             =   945
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Max             =   100
            Min             =   0
            MaxLength       =   0
            Text            =   "0"
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
         Begin UpDonwOCX.UpDown updBrack 
            Height          =   285
            Left            =   1440
            TabIndex        =   58
            Top             =   1305
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Max             =   100
            Min             =   0
            MaxLength       =   0
            Text            =   "0"
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
         Begin VB.Label Label3 
            Caption         =   "Acceleration:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Air Resistance:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Brakebalance:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Track Grip:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Gears"
         Height          =   2726
         Left            =   2355
         TabIndex        =   9
         Top             =   360
         Width           =   2295
         Begin UpDonwOCX.UpDown updGear 
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   70
            Top             =   345
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Value           =   16
            Max             =   75
            Min             =   16
            MaxLength       =   0
            Text            =   "16"
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
         Begin UpDonwOCX.UpDown updGear 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   69
            Top             =   705
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Value           =   17
            Max             =   76
            Min             =   17
            MaxLength       =   0
            Text            =   "17"
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
         Begin UpDonwOCX.UpDown updGear 
            Height          =   285
            Index           =   2
            Left            =   1320
            TabIndex        =   68
            Top             =   1065
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Value           =   18
            Max             =   77
            Min             =   18
            MaxLength       =   0
            Text            =   "18"
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
         Begin UpDonwOCX.UpDown updGear 
            Height          =   285
            Index           =   3
            Left            =   1320
            TabIndex        =   67
            Top             =   1425
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Value           =   19
            Max             =   78
            Min             =   19
            MaxLength       =   0
            Text            =   "19"
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
         Begin UpDonwOCX.UpDown updGear 
            Height          =   285
            Index           =   4
            Left            =   1320
            TabIndex        =   66
            Top             =   1785
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Value           =   20
            Max             =   79
            Min             =   20
            MaxLength       =   0
            Text            =   "20"
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
         Begin UpDonwOCX.UpDown updGear 
            Height          =   285
            Index           =   5
            Left            =   1320
            TabIndex        =   65
            Top             =   2145
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Value           =   21
            Max             =   80
            Min             =   21
            MaxLength       =   0
            Text            =   "21"
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
         Begin VB.Label Label7 
            Caption         =   "1st Gear:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "3rd Gear:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "4th Gear:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "6th Gear:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "5th Gear:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "2nd Gear:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tyre Compound"
         Height          =   735
         Left            =   2355
         TabIndex        =   6
         Top             =   3120
         Width           =   2295
         Begin VB.ComboBox cboTire 
            Height          =   315
            ItemData        =   "frmCCSetup.frx":0342
            Left            =   1440
            List            =   "frmCCSetup.frx":0355
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Tyre Compound:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1170
         End
      End
      Begin VB.Frame fraWing 
         Caption         =   "Wings"
         Height          =   1000
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2175
         Begin UpDonwOCX.UpDown updFront 
            Height          =   285
            Left            =   1440
            TabIndex        =   64
            Top             =   225
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Value           =   1
            Max             =   20
            Min             =   1
            MaxLength       =   2
            Text            =   "1"
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
         Begin UpDonwOCX.UpDown updRear 
            Height          =   285
            Left            =   1440
            TabIndex        =   63
            Top             =   585
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Value           =   1
            Max             =   20
            Min             =   1
            MaxLength       =   0
            Text            =   "1"
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
         Begin VB.Label Label1 
            Caption         =   "Front Wing:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Rear Wing:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Laps"
         Height          =   195
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   4920
      TabIndex        =   0
      Top             =   4280
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   4280
      Width           =   1035
   End
End
Attribute VB_Name = "frmCCSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Path As String
Dim Update As Boolean

Private Sub cmdCancel_Click()
    On Error Resume Next
    Kill ProgramDir & "\File\CCSetup.tmp"
    Kill ProgramDir & "\File\PitStop.tmp"
    Unload Me
End Sub

Private Sub cmdSave_Click()
    frmCCSetup.MousePointer = 11
    SaveCCSetup
    SavePitStop
    WriteCheckSum Path
    frmCCSetup.MousePointer = 0
End Sub

Private Sub Form_Load()
    Update = False
    tabMain.Tab = 0
    If frmMain.tabMain.Tab = 0 Then
        Path = frmMain.lstFile.SelectedItem.Key
    ElseIf frmMain.tabMain.Tab = 1 Then
        Path = frmMain.txtPath.Text
    End If
    LoadPitStop
    LoadCCSetup
End Sub

Private Sub upd1Stops_Change()
    If Update = False Then
        If upd1Stops.Text = "" Then
            cmdSave.Enabled = False
            Exit Sub
        End If
        X = upd1Stops.Text + upd2Stops.Value + upd3Stops.Value
        If X <> 100 Then
            cmdSave.Enabled = False
        Else
            cmdSave.Enabled = True
        End If
    End If
End Sub

Private Sub upd2Stops_Change()
    If Update = False Then
        If upd2Stops.Text = "" Then
            cmdSave.Enabled = False
            Exit Sub
        End If
        X = upd1Stops.Value + upd2Stops.Text + upd3Stops.Value
        If X <> 100 Then
            cmdSave.Enabled = False
        Else
            cmdSave.Enabled = True
        End If
    End If
End Sub

Private Sub upd3Stops_Change()
    If Update = False Then
        If upd3Stops.Text = "" Then
            cmdSave.Enabled = False
            Exit Sub
        End If
        X = upd1Stops.Value + upd2Stops.Value + upd3Stops.Text
        If X <> 100 Then
            cmdSave.Enabled = False
        Else
            cmdSave.Enabled = True
        End If
    End If
End Sub

Private Sub LoadPitStop()
'*************************************
'Function Name: LoadPitStop
'Use: Load the CC Car Pit Stop Strategy
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-05
'*************************************
Dim iInt As Integer
    On Error GoTo ErrHandler
    Update = True
    GetPitStop Path
    FileNum = FreeFile
    Open ProgramDir & "\File\PitStop.tmp" For Binary As FileNum

    'Load Laps
    Get #FileNum, 1, iInt
    updLaps.Value = iInt

    'Load data for One stop
    Get #FileNum, 7, iInt
    upd1Stops.Value = iInt
    Get #FileNum, 9, iInt
    updStart11.Value = iInt
    Get #FileNum, 11, iInt
    updRange11.Value = iInt

    'Load data for two stops
    Get #FileNum, 23, iInt
    upd2Stops.Value = iInt
    Get #FileNum, 25, iInt
    updStart21.Value = iInt
    Get #FileNum, 27, iInt
    updRange21.Value = iInt
    Get #FileNum, 29, iInt
    updStart22.Value = iInt
    Get #FileNum, 31, iInt
    updRange22.Value = iInt

    'Load Data for three stops
    Get #FileNum, 39, iInt
    upd3Stops.Value = iInt
    Get #FileNum, 41, iInt
    updStart31.Value = iInt
    Get #FileNum, 43, iInt
    updRange31.Value = iInt
    Get #FileNum, 45, iInt
    updStart32.Value = iInt
    Get #FileNum, 47, iInt
    updRange32.Value = iInt
    Get #FileNum, 49, iInt
    updStart33.Value = iInt
    Get #FileNum, 51, iInt
    updRange33.Value = iInt
    Close FileNum
    Update = False
Exit Sub
ErrHandler:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: " & Err.Source, vbCritical, "Error"
    Unload Me
End Sub

Private Sub LoadCCSetup()
'*************************************
'Function Name: LoadCCSetup
'Use: Load the CC Car Setup
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-05
'*************************************
Dim bByte As Byte
Dim iInt As Integer
    On Error GoTo ErrHandler
    GetCCSetup Path
    FileNum = FreeFile
    Open ProgramDir & "\File\CCSetup.tmp" For Binary As FileNum
    
    'Load Wings
    Get #FileNum, 2, bByte
    If bByte - 151 > 20 Then updFront.Max = bByte - 151
    If bByte - 151 > 20 Then
        updRear.Max = bByte - 151
        MsgBox "The front wing is set to high, it's " & bByte - 151 & "." & vbLf & "The program will show the wing settings but please change this to 20 of less.", vbInformation
    End If
    If bByte - 151 < 1 Then
        MsgBox "The front wing is set to low, it's " & bByte - 151 & "." & vbLf & "The program will show the wing as it was set to 1.", vbInformation
        bByte = 152
    End If
    updFront.Value = bByte - 151
    Get #FileNum, 3, bByte
    If bByte - 151 > 20 Then
        updRear.Max = bByte - 151
        MsgBox "The rear wing is set to high, it's " & bByte - 151 & "." & vbLf & "The program will show the wing settings but please change this to 20 of less.", vbInformation
    End If
    If bByte - 151 < 1 Then
        MsgBox "The rear wing is set to low, it's " & bByte - 151 & "." & vbLf & "The program will show the wing as it was set to 1.", vbInformation
        bByte = 152
    End If
    updRear.Value = bByte - 151
    
    'Load Gears
    Get #FileNum, 4, bByte
    updGear(0).Value = bByte - 151
    Get #FileNum, 5, bByte
    updGear(1).Value = bByte - 151
    Get #FileNum, 6, bByte
    updGear(2).Value = bByte - 151
    Get #FileNum, 7, bByte
    updGear(3).Value = bByte - 151
    Get #FileNum, 8, bByte
    updGear(4).Value = bByte - 151
    Get #FileNum, 9, bByte
    updGear(5).Value = bByte - 151
    
    'Load Tire
    Get #FileNum, 10, bByte
    If bByte = 103 Then
        cboTire.ListIndex = 4
    Else
        cboTire.ListIndex = bByte - 52
    End If
    
    'Load Psysics
    Get #FileNum, 13, bByte
    updGrip.Value = bByte
    Get #FileNum, 15, bByte
    updBrack.Value = bByte
    Get #FileNum, 21, bByte
    updAcc.Value = bByte
    Get #FileNum, 23, bByte
    updAir.Value = bByte

    'Load Fuel
    Get #FileNum, 28, iInt
    updFuel.Value = iInt

    Close FileNum
Exit Sub
ErrHandler:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: " & Err.Source, vbCritical, "Error"
    Unload Me
End Sub

Private Sub SavePitStop()
'*************************************
'Function Name: SavePitStop
'Use: Save CC CAr Pit Stop Strategy
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-05
'*************************************
Dim iInt As Integer
On Error GoTo ErrHandler
    FileNum = FreeFile
    Open ProgramDir & "\File\PitStop.tmp" For Binary As FileNum

    'Save Laps
    iInt = updLaps.Value
    Put #FileNum, 1, iInt

    'Save data for One stop
    iInt = upd1Stops.Value
    Put #FileNum, 7, iInt
    iInt = updStart11.Value
    Put #FileNum, 9, iInt
    iInt = updRange11.Value
    Put #FileNum, 11, iInt

    'Save data for two stops
    iInt = upd2Stops.Value
    Put #FileNum, 23, iInt
    iInt = updStart21.Value
    Put #FileNum, 25, iInt
    iInt = updRange21.Value
    Put #FileNum, 27, iInt
    iInt = updStart22.Value
    Put #FileNum, 29, iInt
    iInt = updRange22.Value
    Put #FileNum, 31, iInt

    'Save Data for three stops
    iInt = upd3Stops.Value
    Put #FileNum, 39, iInt
    iInt = updStart31.Value
    Put #FileNum, 41, iInt
    iInt = updRange31.Value
    Put #FileNum, 43, iInt
    iInt = updStart32.Value
    Put #FileNum, 45, iInt
    iInt = updRange32.Value
    Put #FileNum, 47, iInt
    iInt = updStart33.Value
    Put #FileNum, 49, iInt
    iInt = updRange33.Value
    Put #FileNum, 51, iInt
    Close FileNum
    SaveCCPitStop Path
Exit Sub
ErrHandler:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: " & Err.Source, vbCritical, "Error"
    Unload Me
End Sub

Private Sub SaveCCSetup()
'*************************************
'Function Name: SaveCCSetup
'Use: Save CC Car Setup
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-05
'*************************************
Dim bByte As Byte
Dim iInt As Integer
On Error GoTo ErrHandler

    GetCCSetup Path
    FileNum = FreeFile
    Open ProgramDir & "\File\CCSetup.tmp" For Binary As FileNum

    'Save Wings
    bByte = updFront.Value + 151
    Put #FileNum, 2, bByte
    bByte = updRear.Value + 151
    Put #FileNum, 3, bByte

    'Save Gears
    bByte = updGear(0).Value + 151
    Put #FileNum, 4, bByte
    bByte = updGear(1).Value + 151
    Put #FileNum, 5, bByte
    bByte = updGear(2).Value + 151
    Put #FileNum, 6, bByte
    bByte = updGear(3).Value + 151
    Put #FileNum, 7, bByte
    bByte = updGear(4).Value + 151
    Put #FileNum, 8, bByte
    bByte = updGear(5).Value + 151
    Put #FileNum, 9, bByte

    'Save Tire
    If cboTire.Text = "Unk" Then
        bByte = 103
    Else
        bByte = cboTire.ListIndex + 52
    End If
    Put #FileNum, 10, bByte

    'Save Psysics
    bByte = updGrip.Value
    Put #FileNum, 13, bByte
    bByte = updBrack.Value
    Put #FileNum, 15, bByte
    bByte = updAcc.Value
    Put #FileNum, 21, bByte
    bByte = updAir.Value
    Put #FileNum, 23, bByte

    'Save Fuel
    iInt = updFuel.Text
    Put #FileNum, 28, iInt

    Close FileNum
    SaveCCCarSetup Path
Exit Sub
ErrHandler:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: " & Err.Source, vbCritical, "Error"
    Unload Me
End Sub
