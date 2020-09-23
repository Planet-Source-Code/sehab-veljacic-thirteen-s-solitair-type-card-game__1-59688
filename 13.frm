VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "13's"
   ClientHeight    =   8205
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   547
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   716
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9030
      Left            =   120
      Picture         =   "13.frx":0000
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   700
      TabIndex        =   0
      Top             =   120
      Width           =   10530
      Begin VB.Timer FTF 
         Interval        =   1
         Left            =   8520
         Top             =   7320
      End
      Begin VB.Frame Frame1 
         Caption         =   "HELP"
         Height          =   6255
         Left            =   600
         TabIndex        =   61
         Top             =   840
         Visible         =   0   'False
         Width           =   9615
         Begin VB.Label Label5 
            Caption         =   "X"
            Height          =   375
            Left            =   9240
            TabIndex        =   63
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   $"13.frx":133A22
            Height          =   5415
            Left            =   360
            TabIndex        =   62
            Top             =   600
            Width           =   9015
         End
      End
      Begin VB.Timer Win 
         Interval        =   1
         Left            =   8760
         Top             =   4560
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   4500
         Picture         =   "13.frx":133BB7
         ScaleHeight     =   1440
         ScaleWidth      =   1035
         TabIndex        =   59
         Top             =   5250
         Width           =   1065
      End
      Begin VB.CommandButton ReDeck 
         Caption         =   "Redeck"
         Height          =   495
         Left            =   3960
         TabIndex        =   57
         Top             =   6840
         Width           =   1095
      End
      Begin VB.PictureBox SC 
         BackColor       =   &H00008080&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   3120
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   56
         Top             =   3600
         Width           =   135
      End
      Begin VB.PictureBox FC 
         BackColor       =   &H00008080&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   2760
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   55
         Top             =   3600
         Width           =   135
      End
      Begin VB.Timer Check 
         Interval        =   1
         Left            =   8280
         Top             =   3600
      End
      Begin VB.Timer DeckPile 
         Interval        =   1
         Left            =   8520
         Top             =   4080
      End
      Begin VB.Timer MoveTimer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   9000
         Top             =   4080
      End
      Begin VB.Timer Score 
         Interval        =   10
         Left            =   8760
         Top             =   3600
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1425
         Index           =   1
         Left            =   4800
         Picture         =   "13.frx":1389F9
         ScaleHeight     =   1425
         ScaleWidth      =   1050
         TabIndex        =   14
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Timer MoveTimer 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   9240
         Top             =   3600
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   13
         Left            =   4920
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   52
         Top             =   7080
         Width           =   615
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   12
         Left            =   6360
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   51
         Top             =   6840
         Width           =   615
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   11
         Left            =   5640
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   50
         Top             =   6840
         Width           =   615
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   10
         Left            =   4920
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   49
         Top             =   6840
         Width           =   615
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   9
         Left            =   6360
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   48
         Top             =   6600
         Width           =   615
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   8
         Left            =   5640
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   47
         Top             =   6600
         Width           =   615
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   7
         Left            =   4920
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   46
         Top             =   6600
         Width           =   615
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   6
         Left            =   6360
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   45
         Top             =   6360
         Width           =   615
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   5
         Left            =   5640
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   44
         Top             =   6360
         Width           =   615
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   4
         Left            =   4920
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   43
         Top             =   6360
         Width           =   615
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   3
         Left            =   6360
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   42
         Top             =   6120
         Width           =   615
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   2
         Left            =   5640
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   41
         Top             =   6120
         Width           =   615
      End
      Begin VB.PictureBox C 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   1
         Left            =   4920
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   40
         Top             =   6120
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   13
         Left            =   2520
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   39
         Top             =   7080
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   12
         Left            =   3960
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   38
         Top             =   6840
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   11
         Left            =   3240
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   37
         Top             =   6840
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   10
         Left            =   2520
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   36
         Top             =   6840
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   9
         Left            =   3960
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   35
         Top             =   6600
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   8
         Left            =   3240
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   34
         Top             =   6600
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   7
         Left            =   2520
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   33
         Top             =   6600
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   6
         Left            =   3960
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   32
         Top             =   6360
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   5
         Left            =   3240
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   31
         Top             =   6360
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   4
         Left            =   2520
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   30
         Top             =   6360
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   3
         Left            =   3960
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   29
         Top             =   6120
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   2
         Left            =   3240
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   28
         Top             =   6120
         Width           =   615
      End
      Begin VB.PictureBox H 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   1
         Left            =   2520
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   27
         Top             =   6120
         Width           =   615
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   13
         Left            =   4920
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   26
         Top             =   5160
         Width           =   615
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   12
         Left            =   6360
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   25
         Top             =   4920
         Width           =   615
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   11
         Left            =   5640
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   24
         Top             =   4920
         Width           =   615
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   10
         Left            =   4920
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   23
         Top             =   4920
         Width           =   615
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   9
         Left            =   6360
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   22
         Top             =   4680
         Width           =   615
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   8
         Left            =   5640
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   21
         Top             =   4680
         Width           =   615
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   7
         Left            =   4920
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   20
         Top             =   4680
         Width           =   615
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   6
         Left            =   6360
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   19
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   5
         Left            =   5640
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   18
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   4
         Left            =   4920
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   17
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   3
         Left            =   6360
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   16
         Top             =   4200
         Width           =   615
      End
      Begin VB.PictureBox D 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   2
         Left            =   5640
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   15
         Top             =   4200
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   13
         Left            =   2520
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   13
         Top             =   5160
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   12
         Left            =   3960
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   12
         Top             =   4920
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   11
         Left            =   3240
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   11
         Top             =   4920
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   10
         Left            =   2520
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   10
         Top             =   4920
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   9
         Left            =   3960
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   4680
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   8
         Left            =   3240
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   8
         Top             =   4680
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   7
         Left            =   2520
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   7
         Top             =   4680
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   6
         Left            =   3960
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   6
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   5
         Left            =   3240
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   4
         Left            =   2520
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   3
         Left            =   3960
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   3
         Top             =   4200
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   2
         Left            =   3240
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   4200
         Width           =   615
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   1
         Left            =   2520
         ScaleHeight     =   705
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label YouWin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "YOU WIN"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2280
         TabIndex        =   60
         Top             =   3120
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Key:      K=13 A-Q 2-J 3-10 4-9 5-8 6-7"
         ForeColor       =   &H00800000&
         Height          =   3195
         Left            =   1320
         TabIndex        =   58
         Top             =   4920
         Width           =   765
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "5555555555"
         Height          =   495
         Left            =   7920
         TabIndex        =   54
         Top             =   5280
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   1920
         TabIndex        =   53
         Top             =   360
         Width           =   975
      End
      Begin ComctlLib.ImageList Clover 
         Left            =   8280
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   70
         ImageHeight     =   95
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   13
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":13D8E7
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1427E5
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1476E3
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":14C5E1
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1514DF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":156261
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":15B3B3
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1602B1
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":165033
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":169F31
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":16EE2F
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":173D2D
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":178C2B
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ComctlLib.ImageList Heart 
         Left            =   7680
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   69
         ImageHeight     =   95
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   13
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":17DB29
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1828AB
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1877A9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":18C6A7
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1915A5
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1964A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":19B3A1
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1A029F
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1A519D
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1A9E4F
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1AED4D
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1B3C4B
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1B8B49
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ComctlLib.ImageList Spade 
         Left            =   6480
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   70
         ImageHeight     =   95
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   13
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1BDA47
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1C2945
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1C7843
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1CC741
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1D163F
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1D6611
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1DB50F
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1E040D
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1E530B
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1EA209
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1EF107
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1F4005
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1F8F03
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ComctlLib.ImageList Diamond 
         Left            =   7080
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   70
         ImageHeight     =   95
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   13
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":1FDE01
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":202CFF
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":207BFD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":20CAFB
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":2119F9
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":2168F7
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":21B7F5
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":2206F3
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":2255F1
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":22A4EF
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":22F271
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":23416F
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "13.frx":23906D
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuAnimate 
         Caption         =   "Animate"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCBI 
         Caption         =   "Change Background Image -- 1"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "HELP"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Ro
    StartLeft As Integer
    Top As Integer
    Left As Integer
    EndLeft As Integer
End Type

Private Type Slo
    Left As Integer
    Top As Integer
    WhichCard As Integer
    WhichNum As Integer
End Type

Dim Row(7) As Ro

Dim CardNumber(52) As Single ' array for card numbers
Dim CardSuit(52) As Integer
Dim CardHeart(13) As Integer
Dim CardSpade(13) As Integer
Dim CardClover(13) As Integer
Dim CardDimond(13) As Integer
Dim intA As Integer
Dim intB As Integer
Dim intRow As Integer
Dim intC As Integer
Dim intD As Integer


Dim intBGI As Integer
Dim blnAnimate As Boolean
Dim intChecking As Integer
Dim intFirstCard As Integer
Dim intSecondCard As Integer
Dim blnFirstCard As Boolean
Dim blnSecondCard As Boolean
Dim intFirstSuit As Integer
Dim intSecondSuit As Integer
'Dim intCCS1 As Integer
'Dim intCCS2 As Integer
Dim intScore As Integer
Dim intThirdCard As Integer
Dim intThirdSuit As Integer
Dim intTS As Integer
Dim blnZorder1 As Boolean
Dim blnZorder2 As Boolean
Dim Slots(28) As Slo

'this sub is called when cardtogonext is bigger than 37 or form_load is reached
Private Function Shuffle()
Dim Temp As Integer
Dim ItemPicked As Integer
Dim Left As Integer
Dim i As Integer
For i = 1 To 52
CardNumber(i) = i ' load vaules into array , cardnumber(1) = 1, etc
Next i

For Left = 52 To 2 Step -1      ' ------- SHUFFLE START
ItemPicked = Int(Rnd * Left) + 1 ' pick a acard from cards left
Temp = CardNumber(Left)         ' get bottom card and put it as temp
CardNumber(Left) = CardNumber(ItemPicked) ' bottom card = random picked card from cards left
CardNumber(ItemPicked) = Temp ' cardpicked  = how many cards left
CardSuit(ItemPicked) = (Rnd * 3) + 1
Next Left
End Function

Private Sub C_Click(Index As Integer)

    If blnFirstCard = True Then
        If blnSecondCard = False Then
            intSecondCard = Index
            blnSecondCard = True
            intSecondSuit = 1
            SC.Visible = True
            SC.ZOrder 0
            SC.Left = C(Index).Left + 30
            SC.Top = C(Index).Top + 50
        End If
    End If
    If blnFirstCard = False Then
        intFirstCard = Index
        blnFirstCard = True
        intFirstSuit = 1
        FC.Visible = True
        FC.ZOrder 0
        FC.Left = C(Index).Left + 30
        FC.Top = C(Index).Top + 50
    End If
End Sub

Private Sub C_DblClick(Index As Integer)
    intThirdSuit = 2
    intThirdCard = Index
    DeckPile.Enabled = True
    SC.Visible = False
    FC.Visible = False
    intFirstCard = 0
    intSecondCard = 0
    blnFirstCard = False
    blnSecondCard = False
End Sub




'This is all of the checking
Private Sub Check_Timer()
Picture2.ZOrder 0

For intChecking = 1 To 13

    RowCheck2 intChecking
    RowCheck3 intChecking
    RowCheck4 intChecking
    RowCheck5 intChecking
    RowCheck6 intChecking
    RowCheck7 intChecking
        
Next intChecking

End Sub

Private Sub RowCheck2(intChecking As Integer)
'ROW 2--------------------------------------

    '2
    If H(Slots(2).WhichNum).Left = 268 And H(Slots(2).WhichNum).Top = 100 Then
            If Slots(1).WhichCard = 1 Then
                H(Slots(1).WhichNum).Enabled = False
            Else
                H(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 2 Then
                S(Slots(1).WhichNum).Enabled = False
            Else
                S(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 3 Then
                C(Slots(1).WhichNum).Enabled = False
            Else
                C(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 4 Then
                D(Slots(1).WhichNum).Enabled = False
            Else
                D(Slots(1).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(2).WhichNum).Left = 268 And S(Slots(2).WhichNum).Top = 100 Then
            If Slots(1).WhichCard = 1 Then
                H(Slots(1).WhichNum).Enabled = False
            Else
                H(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 2 Then
                S(Slots(1).WhichNum).Enabled = False
            Else
                S(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 3 Then
                C(Slots(1).WhichNum).Enabled = False
            Else
                C(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 4 Then
                D(Slots(1).WhichNum).Enabled = False
            Else
                D(Slots(1).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(2).WhichNum).Left = 268 And C(Slots(2).WhichNum).Top = 100 Then
            If Slots(1).WhichCard = 1 Then
                H(Slots(1).WhichNum).Enabled = False
            Else
                H(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 2 Then
                S(Slots(1).WhichNum).Enabled = False
            Else
                S(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 3 Then
                C(Slots(1).WhichNum).Enabled = False
            Else
                C(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 4 Then
                D(Slots(1).WhichNum).Enabled = False
            Else
                D(Slots(1).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(2).WhichNum).Left = 268 And D(Slots(2).WhichNum).Top = 100 Then
            If Slots(1).WhichCard = 1 Then
                H(Slots(1).WhichNum).Enabled = False
            Else
                H(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 2 Then
                S(Slots(1).WhichNum).Enabled = False
            Else
                S(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 3 Then
                C(Slots(1).WhichNum).Enabled = False
            Else
                C(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 4 Then
                D(Slots(1).WhichNum).Enabled = False
            Else
                D(Slots(1).WhichNum).Enabled = True
            End If
        End If
        
        '3
        If H(Slots(3).WhichNum).Left = 340 And H(Slots(3).WhichNum).Top = 100 Then
            If Slots(1).WhichCard = 1 Then
                H(Slots(1).WhichNum).Enabled = False
            Else
                H(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 2 Then
                S(Slots(1).WhichNum).Enabled = False
            Else
                S(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 3 Then
                C(Slots(1).WhichNum).Enabled = False
            Else
                C(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 4 Then
                D(Slots(1).WhichNum).Enabled = False
            Else
                D(Slots(1).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(3).WhichNum).Left = 340 And S(Slots(3).WhichNum).Top = 100 Then
            If Slots(1).WhichCard = 1 Then
                H(Slots(1).WhichNum).Enabled = False
            Else
                H(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 2 Then
                S(Slots(1).WhichNum).Enabled = False
            Else
                S(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 3 Then
                C(Slots(1).WhichNum).Enabled = False
            Else
                C(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 4 Then
                D(Slots(1).WhichNum).Enabled = False
            Else
                D(Slots(1).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(3).WhichNum).Left = 340 And C(Slots(3).WhichNum).Top = 100 Then
            If Slots(1).WhichCard = 1 Then
                H(Slots(1).WhichNum).Enabled = False
            Else
                H(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 2 Then
                S(Slots(1).WhichNum).Enabled = False
            Else
                S(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 3 Then
                C(Slots(1).WhichNum).Enabled = False
            Else
                C(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 4 Then
                D(Slots(1).WhichNum).Enabled = False
            Else
                D(Slots(1).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(3).WhichNum).Left = 340 And D(Slots(3).WhichNum).Top = 100 Then
            If Slots(1).WhichCard = 1 Then
                H(Slots(1).WhichNum).Enabled = False
            Else
                H(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 2 Then
                S(Slots(1).WhichNum).Enabled = False
            Else
                S(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 3 Then
                C(Slots(1).WhichNum).Enabled = False
            Else
                C(Slots(1).WhichNum).Enabled = True
            End If
            If Slots(1).WhichCard = 4 Then
                D(Slots(1).WhichNum).Enabled = False
            Else
                D(Slots(1).WhichNum).Enabled = True
            End If
        End If
        
        
End Sub


Private Sub RowCheck3(intChecking As Integer)
'ROW 3--------------------------------------
        
        '4
        If H(Slots(4).WhichNum).Left = 232 And H(Slots(4).WhichNum).Top = 125 Then
            If Slots(2).WhichCard = 1 Then
                H(Slots(2).WhichNum).Enabled = False
            Else
                H(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 2 Then
                S(Slots(2).WhichNum).Enabled = False
            Else
                S(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 3 Then
                C(Slots(2).WhichNum).Enabled = False
            Else
                C(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 4 Then
                D(Slots(2).WhichNum).Enabled = False
            Else
                D(Slots(2).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(4).WhichNum).Left = 232 And S(Slots(4).WhichNum).Top = 125 Then
            If Slots(2).WhichCard = 1 Then
                H(Slots(2).WhichNum).Enabled = False
            Else
                H(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 2 Then
                S(Slots(2).WhichNum).Enabled = False
            Else
                S(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 3 Then
                C(Slots(2).WhichNum).Enabled = False
            Else
                C(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 4 Then
                D(Slots(2).WhichNum).Enabled = False
            Else
                D(Slots(2).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(4).WhichNum).Left = 232 And C(Slots(4).WhichNum).Top = 125 Then
            If Slots(2).WhichCard = 1 Then
                H(Slots(2).WhichNum).Enabled = False
            Else
                H(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 2 Then
                S(Slots(2).WhichNum).Enabled = False
            Else
                S(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 3 Then
                C(Slots(2).WhichNum).Enabled = False
            Else
                C(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 4 Then
                D(Slots(2).WhichNum).Enabled = False
            Else
                D(Slots(2).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(4).WhichNum).Left = 232 And D(Slots(4).WhichNum).Top = 125 Then
            If Slots(2).WhichCard = 1 Then
                H(Slots(2).WhichNum).Enabled = False
            Else
                H(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 2 Then
                S(Slots(2).WhichNum).Enabled = False
            Else
                S(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 3 Then
                C(Slots(2).WhichNum).Enabled = False
            Else
                C(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 4 Then
                D(Slots(2).WhichNum).Enabled = False
            Else
                D(Slots(2).WhichNum).Enabled = True
            End If
        End If
        
        '5
        If H(Slots(5).WhichNum).Left = 304 And H(Slots(5).WhichNum).Top = 125 Then
            If Slots(2).WhichCard = 1 Then
                H(Slots(2).WhichNum).Enabled = False
            Else
                H(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 1 Then
                H(Slots(3).WhichNum).Enabled = False
            Else
                H(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 2 Then
                S(Slots(2).WhichNum).Enabled = False
            Else
                S(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 2 Then
                S(Slots(3).WhichNum).Enabled = False
            Else
                S(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 3 Then
                C(Slots(2).WhichNum).Enabled = False
            Else
                C(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 3 Then
                C(Slots(3).WhichNum).Enabled = False
            Else
                C(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 4 Then
                D(Slots(2).WhichNum).Enabled = False
            Else
                D(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 4 Then
                D(Slots(3).WhichNum).Enabled = False
            Else
                D(Slots(3).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(5).WhichNum).Left = 304 And S(Slots(5).WhichNum).Top = 125 Then
            If Slots(2).WhichCard = 1 Then
                H(Slots(2).WhichNum).Enabled = False
            Else
                H(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 1 Then
                H(Slots(3).WhichNum).Enabled = False
            Else
                H(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 2 Then
                S(Slots(2).WhichNum).Enabled = False
            Else
                S(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 2 Then
                S(Slots(3).WhichNum).Enabled = False
            Else
                S(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 3 Then
                C(Slots(2).WhichNum).Enabled = False
            Else
                C(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 3 Then
                C(Slots(3).WhichNum).Enabled = False
            Else
                C(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 4 Then
                D(Slots(2).WhichNum).Enabled = False
            Else
                D(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 4 Then
                D(Slots(3).WhichNum).Enabled = False
            Else
                D(Slots(3).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(5).WhichNum).Left = 304 And C(Slots(5).WhichNum).Top = 125 Then
            If Slots(2).WhichCard = 1 Then
                H(Slots(2).WhichNum).Enabled = False
            Else
                H(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 1 Then
                H(Slots(3).WhichNum).Enabled = False
            Else
                H(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 2 Then
                S(Slots(2).WhichNum).Enabled = False
            Else
                S(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 2 Then
                S(Slots(3).WhichNum).Enabled = False
            Else
                S(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 3 Then
                C(Slots(2).WhichNum).Enabled = False
            Else
                C(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 3 Then
                C(Slots(3).WhichNum).Enabled = False
            Else
                C(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 4 Then
                D(Slots(2).WhichNum).Enabled = False
            Else
                D(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 4 Then
                D(Slots(3).WhichNum).Enabled = False
            Else
                D(Slots(3).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(5).WhichNum).Left = 304 And D(Slots(5).WhichNum).Top = 125 Then
            If Slots(2).WhichCard = 1 Then
                H(Slots(2).WhichNum).Enabled = False
            Else
                H(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 1 Then
                H(Slots(3).WhichNum).Enabled = False
            Else
                H(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 2 Then
                S(Slots(2).WhichNum).Enabled = False
            Else
                S(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 2 Then
                S(Slots(3).WhichNum).Enabled = False
            Else
                S(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 3 Then
                C(Slots(2).WhichNum).Enabled = False
            Else
                C(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 3 Then
                C(Slots(3).WhichNum).Enabled = False
            Else
                C(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(2).WhichCard = 4 Then
                D(Slots(2).WhichNum).Enabled = False
            Else
                D(Slots(2).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 4 Then
                D(Slots(3).WhichNum).Enabled = False
            Else
                D(Slots(3).WhichNum).Enabled = True
            End If
        End If
        
        
        '6
        If H(Slots(6).WhichNum).Left = 376 And H(Slots(6).WhichNum).Top = 125 Then
            If Slots(3).WhichCard = 1 Then
                H(Slots(3).WhichNum).Enabled = False
            Else
                H(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 2 Then
                S(Slots(3).WhichNum).Enabled = False
            Else
                S(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 3 Then
                C(Slots(3).WhichNum).Enabled = False
            Else
                C(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 4 Then
                D(Slots(3).WhichNum).Enabled = False
            Else
                D(Slots(3).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(6).WhichNum).Left = 376 And S(Slots(6).WhichNum).Top = 125 Then
            If Slots(3).WhichCard = 1 Then
                H(Slots(3).WhichNum).Enabled = False
            Else
                H(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 2 Then
                S(Slots(3).WhichNum).Enabled = False
            Else
                S(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 3 Then
                C(Slots(3).WhichNum).Enabled = False
            Else
                C(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 4 Then
                D(Slots(3).WhichNum).Enabled = False
            Else
                D(Slots(3).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(6).WhichNum).Left = 376 And C(Slots(6).WhichNum).Top = 125 Then
            If Slots(3).WhichCard = 1 Then
                H(Slots(3).WhichNum).Enabled = False
            Else
                H(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 2 Then
                S(Slots(3).WhichNum).Enabled = False
            Else
                S(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 3 Then
                C(Slots(3).WhichNum).Enabled = False
            Else
                C(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 4 Then
                D(Slots(3).WhichNum).Enabled = False
            Else
                D(Slots(3).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(6).WhichNum).Left = 376 And D(Slots(6).WhichNum).Top = 125 Then
            If Slots(3).WhichCard = 1 Then
                H(Slots(3).WhichNum).Enabled = False
            Else
                H(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 2 Then
                S(Slots(3).WhichNum).Enabled = False
            Else
                S(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 3 Then
                C(Slots(3).WhichNum).Enabled = False
            Else
                C(Slots(3).WhichNum).Enabled = True
            End If
            If Slots(3).WhichCard = 4 Then
                D(Slots(3).WhichNum).Enabled = False
            Else
                D(Slots(3).WhichNum).Enabled = True
            End If
        End If
        
        
        
        
        
        
End Sub


Private Sub RowCheck4(intChecking As Integer)
'ROW 4-----------------------------------------

        '7
        If H(Slots(7).WhichNum).Left = 196 And H(Slots(7).WhichNum).Top = 150 Then
            If Slots(4).WhichCard = 1 Then
                H(Slots(4).WhichNum).Enabled = False
            Else
                H(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 2 Then
                S(Slots(4).WhichNum).Enabled = False
            Else
                S(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 3 Then
                C(Slots(4).WhichNum).Enabled = False
            Else
                C(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 4 Then
                D(Slots(4).WhichNum).Enabled = False
            Else
                D(Slots(4).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(7).WhichNum).Left = 196 And S(Slots(7).WhichNum).Top = 150 Then
            If Slots(4).WhichCard = 1 Then
                H(Slots(4).WhichNum).Enabled = False
            Else
                H(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 2 Then
                S(Slots(4).WhichNum).Enabled = False
            Else
                S(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 3 Then
                C(Slots(4).WhichNum).Enabled = False
            Else
                C(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 4 Then
                D(Slots(4).WhichNum).Enabled = False
            Else
                D(Slots(4).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(7).WhichNum).Left = 196 And C(Slots(7).WhichNum).Top = 150 Then
            If Slots(4).WhichCard = 1 Then
                H(Slots(4).WhichNum).Enabled = False
            Else
                H(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 2 Then
                S(Slots(4).WhichNum).Enabled = False
            Else
                S(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 3 Then
                C(Slots(4).WhichNum).Enabled = False
            Else
                C(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 4 Then
                D(Slots(4).WhichNum).Enabled = False
            Else
                D(Slots(4).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(7).WhichNum).Left = 196 And D(Slots(7).WhichNum).Top = 150 Then
            If Slots(4).WhichCard = 1 Then
                H(Slots(4).WhichNum).Enabled = False
            Else
                H(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 2 Then
                S(Slots(4).WhichNum).Enabled = False
            Else
                S(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 3 Then
                C(Slots(4).WhichNum).Enabled = False
            Else
                C(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 4 Then
                D(Slots(4).WhichNum).Enabled = False
            Else
                D(Slots(4).WhichNum).Enabled = True
            End If
        End If
        
        '8
        If H(Slots(8).WhichNum).Left = 268 And H(Slots(8).WhichNum).Top = 150 Then
            If Slots(4).WhichCard = 1 Then
                H(Slots(4).WhichNum).Enabled = False
            Else
                H(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 1 Then
                H(Slots(5).WhichNum).Enabled = False
            Else
                H(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 2 Then
                S(Slots(4).WhichNum).Enabled = False
            Else
                S(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 2 Then
                S(Slots(5).WhichNum).Enabled = False
            Else
                S(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 3 Then
                C(Slots(4).WhichNum).Enabled = False
            Else
                C(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 3 Then
                C(Slots(5).WhichNum).Enabled = False
            Else
                C(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 4 Then
                D(Slots(4).WhichNum).Enabled = False
            Else
                D(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 4 Then
                D(Slots(5).WhichNum).Enabled = False
            Else
                D(Slots(5).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(8).WhichNum).Left = 268 And S(Slots(8).WhichNum).Top = 150 Then
            If Slots(4).WhichCard = 1 Then
                H(Slots(4).WhichNum).Enabled = False
            Else
                H(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 1 Then
                H(Slots(5).WhichNum).Enabled = False
            Else
                H(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 2 Then
                S(Slots(4).WhichNum).Enabled = False
            Else
                S(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 2 Then
                S(Slots(5).WhichNum).Enabled = False
            Else
                S(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 3 Then
                C(Slots(4).WhichNum).Enabled = False
            Else
                C(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 3 Then
                C(Slots(5).WhichNum).Enabled = False
            Else
                C(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 4 Then
                D(Slots(4).WhichNum).Enabled = False
            Else
                D(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 4 Then
                D(Slots(5).WhichNum).Enabled = False
            Else
                D(Slots(5).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(8).WhichNum).Left = 268 And C(Slots(8).WhichNum).Top = 150 Then
            If Slots(4).WhichCard = 1 Then
                H(Slots(4).WhichNum).Enabled = False
            Else
                H(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 1 Then
                H(Slots(5).WhichNum).Enabled = False
            Else
                H(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 2 Then
                S(Slots(4).WhichNum).Enabled = False
            Else
                S(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 2 Then
                S(Slots(5).WhichNum).Enabled = False
            Else
                S(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 3 Then
                C(Slots(4).WhichNum).Enabled = False
            Else
                C(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 3 Then
                C(Slots(5).WhichNum).Enabled = False
            Else
                C(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 4 Then
                D(Slots(4).WhichNum).Enabled = False
            Else
                D(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 4 Then
                D(Slots(5).WhichNum).Enabled = False
            Else
                D(Slots(5).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(8).WhichNum).Left = 268 And D(Slots(8).WhichNum).Top = 150 Then
            If Slots(4).WhichCard = 1 Then
                H(Slots(4).WhichNum).Enabled = False
            Else
                H(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 1 Then
                H(Slots(5).WhichNum).Enabled = False
            Else
                H(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 2 Then
                S(Slots(4).WhichNum).Enabled = False
            Else
                S(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 2 Then
                S(Slots(5).WhichNum).Enabled = False
            Else
                S(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 3 Then
                C(Slots(4).WhichNum).Enabled = False
            Else
                C(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 3 Then
                C(Slots(5).WhichNum).Enabled = False
            Else
                C(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(4).WhichCard = 4 Then
                D(Slots(4).WhichNum).Enabled = False
            Else
                D(Slots(4).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 4 Then
                D(Slots(5).WhichNum).Enabled = False
            Else
                D(Slots(5).WhichNum).Enabled = True
            End If
        End If
        
        '9
        If H(Slots(9).WhichNum).Left = 340 And H(Slots(9).WhichNum).Top = 150 Then
            If Slots(5).WhichCard = 1 Then
                H(Slots(5).WhichNum).Enabled = False
            Else
                H(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 1 Then
                H(Slots(6).WhichNum).Enabled = False
            Else
                H(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 2 Then
                S(Slots(5).WhichNum).Enabled = False
            Else
                S(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 2 Then
                S(Slots(6).WhichNum).Enabled = False
            Else
                S(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 3 Then
                C(Slots(5).WhichNum).Enabled = False
            Else
                C(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 3 Then
                C(Slots(6).WhichNum).Enabled = False
            Else
                C(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 4 Then
                D(Slots(5).WhichNum).Enabled = False
            Else
                D(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 4 Then
                D(Slots(6).WhichNum).Enabled = False
            Else
                D(Slots(6).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(9).WhichNum).Left = 340 And S(Slots(9).WhichNum).Top = 150 Then
            If Slots(5).WhichCard = 1 Then
                H(Slots(5).WhichNum).Enabled = False
            Else
                H(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 1 Then
                H(Slots(6).WhichNum).Enabled = False
            Else
                H(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 2 Then
                S(Slots(5).WhichNum).Enabled = False
            Else
                S(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 2 Then
                S(Slots(6).WhichNum).Enabled = False
            Else
                S(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 3 Then
                C(Slots(5).WhichNum).Enabled = False
            Else
                C(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 3 Then
                C(Slots(6).WhichNum).Enabled = False
            Else
                C(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 4 Then
                D(Slots(5).WhichNum).Enabled = False
            Else
                D(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 4 Then
                D(Slots(6).WhichNum).Enabled = False
            Else
                D(Slots(6).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(9).WhichNum).Left = 340 And C(Slots(9).WhichNum).Top = 150 Then
            If Slots(5).WhichCard = 1 Then
                H(Slots(5).WhichNum).Enabled = False
            Else
                H(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 1 Then
                H(Slots(6).WhichNum).Enabled = False
            Else
                H(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 2 Then
                S(Slots(5).WhichNum).Enabled = False
            Else
                S(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 2 Then
                S(Slots(6).WhichNum).Enabled = False
            Else
                S(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 3 Then
                C(Slots(5).WhichNum).Enabled = False
            Else
                C(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 3 Then
                C(Slots(6).WhichNum).Enabled = False
            Else
                C(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 4 Then
                D(Slots(5).WhichNum).Enabled = False
            Else
                D(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 4 Then
                D(Slots(6).WhichNum).Enabled = False
            Else
                D(Slots(6).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(9).WhichNum).Left = 340 And D(Slots(9).WhichNum).Top = 150 Then
            If Slots(5).WhichCard = 1 Then
                H(Slots(5).WhichNum).Enabled = False
            Else
                H(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 1 Then
                H(Slots(6).WhichNum).Enabled = False
            Else
                H(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 2 Then
                S(Slots(5).WhichNum).Enabled = False
            Else
                S(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 2 Then
                S(Slots(6).WhichNum).Enabled = False
            Else
                S(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 3 Then
                C(Slots(5).WhichNum).Enabled = False
            Else
                C(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 3 Then
                C(Slots(6).WhichNum).Enabled = False
            Else
                C(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(5).WhichCard = 4 Then
                D(Slots(5).WhichNum).Enabled = False
            Else
                D(Slots(5).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 4 Then
                D(Slots(6).WhichNum).Enabled = False
            Else
                D(Slots(6).WhichNum).Enabled = True
            End If
        End If
        
        '10
        If H(Slots(10).WhichNum).Left = 412 And H(Slots(10).WhichNum).Top = 150 Then
            If Slots(6).WhichCard = 1 Then
                H(Slots(6).WhichNum).Enabled = False
            Else
                H(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 2 Then
                S(Slots(6).WhichNum).Enabled = False
            Else
                S(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 3 Then
                C(Slots(6).WhichNum).Enabled = False
            Else
                C(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 4 Then
                D(Slots(6).WhichNum).Enabled = False
            Else
                D(Slots(6).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(10).WhichNum).Left = 412 And S(Slots(10).WhichNum).Top = 150 Then
            If Slots(6).WhichCard = 1 Then
                H(Slots(6).WhichNum).Enabled = False
            Else
                H(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 2 Then
                S(Slots(6).WhichNum).Enabled = False
            Else
                S(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 3 Then
                C(Slots(6).WhichNum).Enabled = False
            Else
                C(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 4 Then
                D(Slots(6).WhichNum).Enabled = False
            Else
                D(Slots(6).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(10).WhichNum).Left = 412 And C(Slots(10).WhichNum).Top = 150 Then
            If Slots(6).WhichCard = 1 Then
                H(Slots(6).WhichNum).Enabled = False
            Else
                H(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 2 Then
                S(Slots(6).WhichNum).Enabled = False
            Else
                S(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 3 Then
                C(Slots(6).WhichNum).Enabled = False
            Else
                C(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 4 Then
                D(Slots(6).WhichNum).Enabled = False
            Else
                D(Slots(6).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(10).WhichNum).Left = 412 And D(Slots(10).WhichNum).Top = 150 Then
            If Slots(6).WhichCard = 1 Then
                H(Slots(6).WhichNum).Enabled = False
            Else
                H(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 2 Then
                S(Slots(6).WhichNum).Enabled = False
            Else
                S(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 3 Then
                C(Slots(6).WhichNum).Enabled = False
            Else
                C(Slots(6).WhichNum).Enabled = True
            End If
            If Slots(6).WhichCard = 4 Then
                D(Slots(6).WhichNum).Enabled = False
            Else
                D(Slots(6).WhichNum).Enabled = True
            End If
        End If
        
        
        
        
End Sub

Private Sub RowCheck5(intChecking As Integer)
'ROW 5------------------------------------
        
        
        '11
        If H(Slots(11).WhichNum).Left = 160 And H(Slots(11).WhichNum).Top = 175 Then
            If Slots(7).WhichCard = 1 Then
                H(Slots(7).WhichNum).Enabled = False
            Else
                H(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 2 Then
                S(Slots(7).WhichNum).Enabled = False
            Else
                S(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 3 Then
                C(Slots(7).WhichNum).Enabled = False
            Else
                C(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 4 Then
                D(Slots(7).WhichNum).Enabled = False
            Else
                D(Slots(7).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(11).WhichNum).Left = 160 And S(Slots(11).WhichNum).Top = 175 Then
            If Slots(7).WhichCard = 1 Then
                H(Slots(7).WhichNum).Enabled = False
            Else
                H(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 2 Then
                S(Slots(7).WhichNum).Enabled = False
            Else
                S(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 3 Then
                C(Slots(7).WhichNum).Enabled = False
            Else
                C(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 4 Then
                D(Slots(7).WhichNum).Enabled = False
            Else
                D(Slots(7).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(11).WhichNum).Left = 160 And C(Slots(11).WhichNum).Top = 175 Then
            If Slots(7).WhichCard = 1 Then
                H(Slots(7).WhichNum).Enabled = False
            Else
                H(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 2 Then
                S(Slots(7).WhichNum).Enabled = False
            Else
                S(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 3 Then
                C(Slots(7).WhichNum).Enabled = False
            Else
                C(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 4 Then
                D(Slots(7).WhichNum).Enabled = False
            Else
                D(Slots(7).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(11).WhichNum).Left = 160 And D(Slots(11).WhichNum).Top = 175 Then
            If Slots(7).WhichCard = 1 Then
                H(Slots(7).WhichNum).Enabled = False
            Else
                H(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 2 Then
                S(Slots(7).WhichNum).Enabled = False
            Else
                S(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 3 Then
                C(Slots(7).WhichNum).Enabled = False
            Else
                C(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 4 Then
                D(Slots(7).WhichNum).Enabled = False
            Else
                D(Slots(7).WhichNum).Enabled = True
            End If
        End If
        
        '12
        If H(Slots(12).WhichNum).Left = 232 And H(Slots(12).WhichNum).Top = 175 Then
            If Slots(7).WhichCard = 1 Then
                H(Slots(7).WhichNum).Enabled = False
            Else
                H(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 1 Then
                H(Slots(8).WhichNum).Enabled = False
            Else
                H(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 2 Then
                S(Slots(7).WhichNum).Enabled = False
            Else
                S(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 2 Then
                S(Slots(8).WhichNum).Enabled = False
            Else
                S(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 3 Then
                C(Slots(7).WhichNum).Enabled = False
            Else
                C(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 3 Then
                C(Slots(8).WhichNum).Enabled = False
            Else
                C(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 4 Then
                D(Slots(7).WhichNum).Enabled = False
            Else
                D(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 4 Then
                D(Slots(8).WhichNum).Enabled = False
            Else
                D(Slots(8).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(12).WhichNum).Left = 232 And S(Slots(12).WhichNum).Top = 175 Then
            If Slots(7).WhichCard = 1 Then
                H(Slots(7).WhichNum).Enabled = False
            Else
                H(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 1 Then
                H(Slots(8).WhichNum).Enabled = False
            Else
                H(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 2 Then
                S(Slots(7).WhichNum).Enabled = False
            Else
                S(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 2 Then
                S(Slots(8).WhichNum).Enabled = False
            Else
                S(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 3 Then
                C(Slots(7).WhichNum).Enabled = False
            Else
                C(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 3 Then
                C(Slots(8).WhichNum).Enabled = False
            Else
                C(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 4 Then
                D(Slots(7).WhichNum).Enabled = False
            Else
                D(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 4 Then
                D(Slots(8).WhichNum).Enabled = False
            Else
                D(Slots(8).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(12).WhichNum).Left = 232 And C(Slots(12).WhichNum).Top = 175 Then
            If Slots(7).WhichCard = 1 Then
                H(Slots(7).WhichNum).Enabled = False
            Else
                H(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 1 Then
                H(Slots(8).WhichNum).Enabled = False
            Else
                H(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 2 Then
                S(Slots(7).WhichNum).Enabled = False
            Else
                S(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 2 Then
                S(Slots(8).WhichNum).Enabled = False
            Else
                S(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 3 Then
                C(Slots(7).WhichNum).Enabled = False
            Else
                C(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 3 Then
                C(Slots(8).WhichNum).Enabled = False
            Else
                C(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 4 Then
                D(Slots(7).WhichNum).Enabled = False
            Else
                D(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 4 Then
                D(Slots(8).WhichNum).Enabled = False
            Else
                D(Slots(8).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(12).WhichNum).Left = 232 And D(Slots(12).WhichNum).Top = 175 Then
            If Slots(7).WhichCard = 1 Then
                H(Slots(7).WhichNum).Enabled = False
            Else
                H(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 1 Then
                H(Slots(8).WhichNum).Enabled = False
            Else
                H(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 2 Then
                S(Slots(7).WhichNum).Enabled = False
            Else
                S(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 2 Then
                S(Slots(8).WhichNum).Enabled = False
            Else
                S(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 3 Then
                C(Slots(7).WhichNum).Enabled = False
            Else
                C(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 3 Then
                C(Slots(8).WhichNum).Enabled = False
            Else
                C(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(7).WhichCard = 4 Then
                D(Slots(7).WhichNum).Enabled = False
            Else
                D(Slots(7).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 4 Then
                D(Slots(8).WhichNum).Enabled = False
            Else
                D(Slots(8).WhichNum).Enabled = True
            End If
        End If
        
        
        '13
        If H(Slots(13).WhichNum).Left = 304 And H(Slots(13).WhichNum).Top = 175 Then
            If Slots(8).WhichCard = 1 Then
                H(Slots(8).WhichNum).Enabled = False
            Else
                H(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 1 Then
                H(Slots(9).WhichNum).Enabled = False
            Else
                H(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 2 Then
                S(Slots(8).WhichNum).Enabled = False
            Else
                S(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 2 Then
                S(Slots(9).WhichNum).Enabled = False
            Else
                S(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 3 Then
                C(Slots(8).WhichNum).Enabled = False
            Else
                C(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 3 Then
                C(Slots(9).WhichNum).Enabled = False
            Else
                C(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 4 Then
                D(Slots(8).WhichNum).Enabled = False
            Else
                D(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 4 Then
                D(Slots(9).WhichNum).Enabled = False
            Else
                D(Slots(9).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(13).WhichNum).Left = 304 And S(Slots(13).WhichNum).Top = 175 Then
            If Slots(8).WhichCard = 1 Then
                H(Slots(8).WhichNum).Enabled = False
            Else
                H(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 1 Then
                H(Slots(9).WhichNum).Enabled = False
            Else
                H(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 2 Then
                S(Slots(8).WhichNum).Enabled = False
            Else
                S(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 2 Then
                S(Slots(9).WhichNum).Enabled = False
            Else
                S(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 3 Then
                C(Slots(8).WhichNum).Enabled = False
            Else
                C(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 3 Then
                C(Slots(9).WhichNum).Enabled = False
            Else
                C(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 4 Then
                D(Slots(8).WhichNum).Enabled = False
            Else
                D(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 4 Then
                D(Slots(9).WhichNum).Enabled = False
            Else
                D(Slots(9).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(13).WhichNum).Left = 304 And C(Slots(13).WhichNum).Top = 175 Then
            If Slots(8).WhichCard = 1 Then
                H(Slots(8).WhichNum).Enabled = False
            Else
                H(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 1 Then
                H(Slots(9).WhichNum).Enabled = False
            Else
                H(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 2 Then
                S(Slots(8).WhichNum).Enabled = False
            Else
                S(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 2 Then
                S(Slots(9).WhichNum).Enabled = False
            Else
                S(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 3 Then
                C(Slots(8).WhichNum).Enabled = False
            Else
                C(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 3 Then
                C(Slots(9).WhichNum).Enabled = False
            Else
                C(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 4 Then
                D(Slots(8).WhichNum).Enabled = False
            Else
                D(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 4 Then
                D(Slots(9).WhichNum).Enabled = False
            Else
                D(Slots(9).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(13).WhichNum).Left = 304 And D(Slots(13).WhichNum).Top = 175 Then
            If Slots(8).WhichCard = 1 Then
                H(Slots(8).WhichNum).Enabled = False
            Else
                H(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 1 Then
                H(Slots(9).WhichNum).Enabled = False
            Else
                H(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 2 Then
                S(Slots(8).WhichNum).Enabled = False
            Else
                S(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 2 Then
                S(Slots(9).WhichNum).Enabled = False
            Else
                S(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 3 Then
                C(Slots(8).WhichNum).Enabled = False
            Else
                C(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 3 Then
                C(Slots(9).WhichNum).Enabled = False
            Else
                C(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(8).WhichCard = 4 Then
                D(Slots(8).WhichNum).Enabled = False
            Else
                D(Slots(8).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 4 Then
                D(Slots(9).WhichNum).Enabled = False
            Else
                D(Slots(9).WhichNum).Enabled = True
            End If
        End If
        
        
        '14
        If H(Slots(14).WhichNum).Left = 376 And H(Slots(14).WhichNum).Top = 175 Then
            If Slots(9).WhichCard = 1 Then
                H(Slots(9).WhichNum).Enabled = False
            Else
                H(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 1 Then
                H(Slots(10).WhichNum).Enabled = False
            Else
                H(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 2 Then
                S(Slots(9).WhichNum).Enabled = False
            Else
                S(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 2 Then
                S(Slots(10).WhichNum).Enabled = False
            Else
                S(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 3 Then
                C(Slots(9).WhichNum).Enabled = False
            Else
                C(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 3 Then
                C(Slots(10).WhichNum).Enabled = False
            Else
                C(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 4 Then
                D(Slots(9).WhichNum).Enabled = False
            Else
                D(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 4 Then
                D(Slots(10).WhichNum).Enabled = False
            Else
                D(Slots(10).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(14).WhichNum).Left = 376 And S(Slots(14).WhichNum).Top = 175 Then
            If Slots(9).WhichCard = 1 Then
                H(Slots(9).WhichNum).Enabled = False
            Else
                H(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 1 Then
                H(Slots(10).WhichNum).Enabled = False
            Else
                H(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 2 Then
                S(Slots(9).WhichNum).Enabled = False
            Else
                S(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 2 Then
                S(Slots(10).WhichNum).Enabled = False
            Else
                S(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 3 Then
                C(Slots(9).WhichNum).Enabled = False
            Else
                C(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 3 Then
                C(Slots(10).WhichNum).Enabled = False
            Else
                C(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 4 Then
                D(Slots(9).WhichNum).Enabled = False
            Else
                D(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 4 Then
                D(Slots(10).WhichNum).Enabled = False
            Else
                D(Slots(10).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(14).WhichNum).Left = 376 And C(Slots(14).WhichNum).Top = 175 Then
            If Slots(9).WhichCard = 1 Then
                H(Slots(9).WhichNum).Enabled = False
            Else
                H(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 1 Then
                H(Slots(10).WhichNum).Enabled = False
            Else
                H(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 2 Then
                S(Slots(9).WhichNum).Enabled = False
            Else
                S(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 2 Then
                S(Slots(10).WhichNum).Enabled = False
            Else
                S(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 3 Then
                C(Slots(9).WhichNum).Enabled = False
            Else
                C(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 3 Then
                C(Slots(10).WhichNum).Enabled = False
            Else
                C(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 4 Then
                D(Slots(9).WhichNum).Enabled = False
            Else
                D(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 4 Then
                D(Slots(10).WhichNum).Enabled = False
            Else
                D(Slots(10).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(14).WhichNum).Left = 376 And D(Slots(14).WhichNum).Top = 175 Then
            If Slots(9).WhichCard = 1 Then
                H(Slots(9).WhichNum).Enabled = False
            Else
                H(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 1 Then
                H(Slots(10).WhichNum).Enabled = False
            Else
                H(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 2 Then
                S(Slots(9).WhichNum).Enabled = False
            Else
                S(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 2 Then
                S(Slots(10).WhichNum).Enabled = False
            Else
                S(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 3 Then
                C(Slots(9).WhichNum).Enabled = False
            Else
                C(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 3 Then
                C(Slots(10).WhichNum).Enabled = False
            Else
                C(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(9).WhichCard = 4 Then
                D(Slots(9).WhichNum).Enabled = False
            Else
                D(Slots(9).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 4 Then
                D(Slots(10).WhichNum).Enabled = False
            Else
                D(Slots(10).WhichNum).Enabled = True
            End If
        End If
        
        '15
        If H(Slots(15).WhichNum).Left = 448 And H(Slots(15).WhichNum).Top = 175 Then
            If Slots(10).WhichCard = 1 Then
                H(Slots(10).WhichNum).Enabled = False
            Else
                H(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 2 Then
                S(Slots(10).WhichNum).Enabled = False
            Else
                S(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 3 Then
                C(Slots(10).WhichNum).Enabled = False
            Else
                C(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 4 Then
                D(Slots(10).WhichNum).Enabled = False
            Else
                D(Slots(10).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(15).WhichNum).Left = 448 And S(Slots(15).WhichNum).Top = 175 Then
            If Slots(10).WhichCard = 1 Then
                H(Slots(10).WhichNum).Enabled = False
            Else
                H(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 2 Then
                S(Slots(10).WhichNum).Enabled = False
            Else
                S(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 3 Then
                C(Slots(10).WhichNum).Enabled = False
            Else
                C(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 4 Then
                D(Slots(10).WhichNum).Enabled = False
            Else
                D(Slots(10).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(15).WhichNum).Left = 448 And C(Slots(15).WhichNum).Top = 175 Then
            If Slots(10).WhichCard = 1 Then
                H(Slots(10).WhichNum).Enabled = False
            Else
                H(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 2 Then
                S(Slots(10).WhichNum).Enabled = False
            Else
                S(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 3 Then
                C(Slots(10).WhichNum).Enabled = False
            Else
                C(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 4 Then
                D(Slots(10).WhichNum).Enabled = False
            Else
                D(Slots(10).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(15).WhichNum).Left = 448 And D(Slots(15).WhichNum).Top = 175 Then
            If Slots(10).WhichCard = 1 Then
                H(Slots(10).WhichNum).Enabled = False
            Else
                H(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 2 Then
                S(Slots(10).WhichNum).Enabled = False
            Else
                S(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 3 Then
                C(Slots(10).WhichNum).Enabled = False
            Else
                C(Slots(10).WhichNum).Enabled = True
            End If
            If Slots(10).WhichCard = 4 Then
                D(Slots(10).WhichNum).Enabled = False
            Else
                D(Slots(10).WhichNum).Enabled = True
            End If
        End If
End Sub

Private Sub RowCheck6(intChecking As Integer)
'ROW 6------------------------------------
    
        '16
        If H(Slots(16).WhichNum).Left = 124 And H(Slots(16).WhichNum).Top = 200 Then
            If Slots(11).WhichCard = 1 Then
                H(Slots(11).WhichNum).Enabled = False
            Else
                H(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 2 Then
                S(Slots(11).WhichNum).Enabled = False
            Else
                S(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 3 Then
                C(Slots(11).WhichNum).Enabled = False
            Else
                C(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 4 Then
                D(Slots(11).WhichNum).Enabled = False
            Else
                D(Slots(11).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(16).WhichNum).Left = 124 And S(Slots(16).WhichNum).Top = 200 Then
            If Slots(11).WhichCard = 1 Then
                H(Slots(11).WhichNum).Enabled = False
            Else
                H(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 2 Then
                S(Slots(11).WhichNum).Enabled = False
            Else
                S(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 3 Then
                C(Slots(11).WhichNum).Enabled = False
            Else
                C(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 4 Then
                D(Slots(11).WhichNum).Enabled = False
            Else
                D(Slots(11).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(16).WhichNum).Left = 124 And C(Slots(16).WhichNum).Top = 200 Then
            If Slots(11).WhichCard = 1 Then
                H(Slots(11).WhichNum).Enabled = False
            Else
                H(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 2 Then
                S(Slots(11).WhichNum).Enabled = False
            Else
                S(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 3 Then
                C(Slots(11).WhichNum).Enabled = False
            Else
                C(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 4 Then
                D(Slots(11).WhichNum).Enabled = False
            Else
                D(Slots(11).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(16).WhichNum).Left = 124 And D(Slots(16).WhichNum).Top = 200 Then
            If Slots(11).WhichCard = 1 Then
                H(Slots(11).WhichNum).Enabled = False
            Else
                H(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 2 Then
                S(Slots(11).WhichNum).Enabled = False
            Else
                S(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 3 Then
                C(Slots(11).WhichNum).Enabled = False
            Else
                C(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 4 Then
                D(Slots(11).WhichNum).Enabled = False
            Else
                D(Slots(11).WhichNum).Enabled = True
            End If
        End If
        
        '17
        If H(Slots(17).WhichNum).Left = 196 And H(Slots(17).WhichNum).Top = 200 Then
            If Slots(11).WhichCard = 1 Then
                H(Slots(11).WhichNum).Enabled = False
            Else
                H(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 1 Then
                H(Slots(12).WhichNum).Enabled = False
            Else
                H(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 2 Then
                S(Slots(11).WhichNum).Enabled = False
            Else
                S(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 2 Then
                S(Slots(12).WhichNum).Enabled = False
            Else
                S(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 3 Then
                C(Slots(11).WhichNum).Enabled = False
            Else
                C(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 3 Then
                C(Slots(12).WhichNum).Enabled = False
            Else
                C(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 4 Then
                D(Slots(11).WhichNum).Enabled = False
            Else
                D(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 4 Then
                D(Slots(12).WhichNum).Enabled = False
            Else
                D(Slots(12).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(17).WhichNum).Left = 196 And S(Slots(17).WhichNum).Top = 200 Then
            If Slots(11).WhichCard = 1 Then
                H(Slots(11).WhichNum).Enabled = False
            Else
                H(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 1 Then
                H(Slots(12).WhichNum).Enabled = False
            Else
                H(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 2 Then
                S(Slots(11).WhichNum).Enabled = False
            Else
                S(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 2 Then
                S(Slots(12).WhichNum).Enabled = False
            Else
                S(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 3 Then
                C(Slots(11).WhichNum).Enabled = False
            Else
                C(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 3 Then
                C(Slots(12).WhichNum).Enabled = False
            Else
                C(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 4 Then
                D(Slots(11).WhichNum).Enabled = False
            Else
                D(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 4 Then
                D(Slots(12).WhichNum).Enabled = False
            Else
                D(Slots(12).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(17).WhichNum).Left = 196 And C(Slots(17).WhichNum).Top = 200 Then
            If Slots(11).WhichCard = 1 Then
                H(Slots(11).WhichNum).Enabled = False
            Else
                H(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 1 Then
                H(Slots(12).WhichNum).Enabled = False
            Else
                H(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 2 Then
                S(Slots(11).WhichNum).Enabled = False
            Else
                S(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 2 Then
                S(Slots(12).WhichNum).Enabled = False
            Else
                S(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 3 Then
                C(Slots(11).WhichNum).Enabled = False
            Else
                C(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 3 Then
                C(Slots(12).WhichNum).Enabled = False
            Else
                C(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 4 Then
                D(Slots(11).WhichNum).Enabled = False
            Else
                D(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 4 Then
                D(Slots(12).WhichNum).Enabled = False
            Else
                D(Slots(12).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(17).WhichNum).Left = 196 And D(Slots(17).WhichNum).Top = 200 Then
            If Slots(11).WhichCard = 1 Then
                H(Slots(11).WhichNum).Enabled = False
            Else
                H(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 1 Then
                H(Slots(12).WhichNum).Enabled = False
            Else
                H(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 2 Then
                S(Slots(11).WhichNum).Enabled = False
            Else
                S(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 2 Then
                S(Slots(12).WhichNum).Enabled = False
            Else
                S(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 3 Then
                C(Slots(11).WhichNum).Enabled = False
            Else
                C(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 3 Then
                C(Slots(12).WhichNum).Enabled = False
            Else
                C(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(11).WhichCard = 4 Then
                D(Slots(11).WhichNum).Enabled = False
            Else
                D(Slots(11).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 4 Then
                D(Slots(12).WhichNum).Enabled = False
            Else
                D(Slots(12).WhichNum).Enabled = True
            End If
        End If
        
        '18
        If H(Slots(18).WhichNum).Left = 268 And H(Slots(18).WhichNum).Top = 200 Then
            If Slots(12).WhichCard = 1 Then
                H(Slots(12).WhichNum).Enabled = False
            Else
                H(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 1 Then
                H(Slots(13).WhichNum).Enabled = False
            Else
                H(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 2 Then
                S(Slots(12).WhichNum).Enabled = False
            Else
                S(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 2 Then
                S(Slots(13).WhichNum).Enabled = False
            Else
                S(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 3 Then
                C(Slots(12).WhichNum).Enabled = False
            Else
                C(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 3 Then
                C(Slots(13).WhichNum).Enabled = False
            Else
                C(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 4 Then
                D(Slots(12).WhichNum).Enabled = False
            Else
                D(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 4 Then
                D(Slots(13).WhichNum).Enabled = False
            Else
                D(Slots(13).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(18).WhichNum).Left = 268 And S(Slots(18).WhichNum).Top = 200 Then
            If Slots(12).WhichCard = 1 Then
                H(Slots(12).WhichNum).Enabled = False
            Else
                H(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 1 Then
                H(Slots(13).WhichNum).Enabled = False
            Else
                H(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 2 Then
                S(Slots(12).WhichNum).Enabled = False
            Else
                S(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 2 Then
                S(Slots(13).WhichNum).Enabled = False
            Else
                S(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 3 Then
                C(Slots(12).WhichNum).Enabled = False
            Else
                C(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 3 Then
                C(Slots(13).WhichNum).Enabled = False
            Else
                C(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 4 Then
                D(Slots(12).WhichNum).Enabled = False
            Else
                D(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 4 Then
                D(Slots(13).WhichNum).Enabled = False
            Else
                D(Slots(13).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(18).WhichNum).Left = 268 And C(Slots(18).WhichNum).Top = 200 Then
            If Slots(12).WhichCard = 1 Then
                H(Slots(12).WhichNum).Enabled = False
            Else
                H(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 1 Then
                H(Slots(13).WhichNum).Enabled = False
            Else
                H(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 2 Then
                S(Slots(12).WhichNum).Enabled = False
            Else
                S(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 2 Then
                S(Slots(13).WhichNum).Enabled = False
            Else
                S(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 3 Then
                C(Slots(12).WhichNum).Enabled = False
            Else
                C(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 3 Then
                C(Slots(13).WhichNum).Enabled = False
            Else
                C(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 4 Then
                D(Slots(12).WhichNum).Enabled = False
            Else
                D(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 4 Then
                D(Slots(13).WhichNum).Enabled = False
            Else
                D(Slots(13).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(18).WhichNum).Left = 268 And D(Slots(18).WhichNum).Top = 200 Then
            If Slots(12).WhichCard = 1 Then
                H(Slots(12).WhichNum).Enabled = False
            Else
                H(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 1 Then
                H(Slots(13).WhichNum).Enabled = False
            Else
                H(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 2 Then
                S(Slots(12).WhichNum).Enabled = False
            Else
                S(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 2 Then
                S(Slots(13).WhichNum).Enabled = False
            Else
                S(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 3 Then
                C(Slots(12).WhichNum).Enabled = False
            Else
                C(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 3 Then
                C(Slots(13).WhichNum).Enabled = False
            Else
                C(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(12).WhichCard = 4 Then
                D(Slots(12).WhichNum).Enabled = False
            Else
                D(Slots(12).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 4 Then
                D(Slots(13).WhichNum).Enabled = False
            Else
                D(Slots(13).WhichNum).Enabled = True
            End If
        End If
        
        '19
        If H(Slots(19).WhichNum).Left = 340 And H(Slots(19).WhichNum).Top = 200 Then
            If Slots(13).WhichCard = 1 Then
                H(Slots(13).WhichNum).Enabled = False
            Else
                H(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 1 Then
                H(Slots(14).WhichNum).Enabled = False
            Else
                H(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 2 Then
                S(Slots(13).WhichNum).Enabled = False
            Else
                S(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 2 Then
                S(Slots(14).WhichNum).Enabled = False
            Else
                S(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 3 Then
                C(Slots(13).WhichNum).Enabled = False
            Else
                C(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 3 Then
                C(Slots(14).WhichNum).Enabled = False
            Else
                C(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 4 Then
                D(Slots(13).WhichNum).Enabled = False
            Else
                D(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 4 Then
                D(Slots(14).WhichNum).Enabled = False
            Else
                D(Slots(14).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(19).WhichNum).Left = 340 And S(Slots(19).WhichNum).Top = 200 Then
            If Slots(13).WhichCard = 1 Then
                H(Slots(13).WhichNum).Enabled = False
            Else
                H(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 1 Then
                H(Slots(14).WhichNum).Enabled = False
            Else
                H(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 2 Then
                S(Slots(13).WhichNum).Enabled = False
            Else
                S(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 2 Then
                S(Slots(14).WhichNum).Enabled = False
            Else
                S(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 3 Then
                C(Slots(13).WhichNum).Enabled = False
            Else
                C(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 3 Then
                C(Slots(14).WhichNum).Enabled = False
            Else
                C(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 4 Then
                D(Slots(13).WhichNum).Enabled = False
            Else
                D(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 4 Then
                D(Slots(14).WhichNum).Enabled = False
            Else
                D(Slots(14).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(19).WhichNum).Left = 340 And C(Slots(19).WhichNum).Top = 200 Then
            If Slots(13).WhichCard = 1 Then
                H(Slots(13).WhichNum).Enabled = False
            Else
                H(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 1 Then
                H(Slots(14).WhichNum).Enabled = False
            Else
                H(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 2 Then
                S(Slots(13).WhichNum).Enabled = False
            Else
                S(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 2 Then
                S(Slots(14).WhichNum).Enabled = False
            Else
                S(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 3 Then
                C(Slots(13).WhichNum).Enabled = False
            Else
                C(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 3 Then
                C(Slots(14).WhichNum).Enabled = False
            Else
                C(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 4 Then
                D(Slots(13).WhichNum).Enabled = False
            Else
                D(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 4 Then
                D(Slots(14).WhichNum).Enabled = False
            Else
                D(Slots(14).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(19).WhichNum).Left = 340 And D(Slots(19).WhichNum).Top = 200 Then
            If Slots(13).WhichCard = 1 Then
                H(Slots(13).WhichNum).Enabled = False
            Else
                H(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 1 Then
                H(Slots(14).WhichNum).Enabled = False
            Else
                H(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 2 Then
                S(Slots(13).WhichNum).Enabled = False
            Else
                S(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 2 Then
                S(Slots(14).WhichNum).Enabled = False
            Else
                S(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 3 Then
                C(Slots(13).WhichNum).Enabled = False
            Else
                C(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 3 Then
                C(Slots(14).WhichNum).Enabled = False
            Else
                C(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(13).WhichCard = 4 Then
                D(Slots(13).WhichNum).Enabled = False
            Else
                D(Slots(13).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 4 Then
                D(Slots(14).WhichNum).Enabled = False
            Else
                D(Slots(14).WhichNum).Enabled = True
            End If
        End If
        
        
        '20
        If H(Slots(20).WhichNum).Left = 412 And H(Slots(20).WhichNum).Top = 200 Then
            If Slots(14).WhichCard = 1 Then
                H(Slots(14).WhichNum).Enabled = False
            Else
                H(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 1 Then
                H(Slots(15).WhichNum).Enabled = False
            Else
                H(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 2 Then
                S(Slots(14).WhichNum).Enabled = False
            Else
                S(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 2 Then
                S(Slots(15).WhichNum).Enabled = False
            Else
                S(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 3 Then
                C(Slots(14).WhichNum).Enabled = False
            Else
                C(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 3 Then
                C(Slots(15).WhichNum).Enabled = False
            Else
                C(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 4 Then
                D(Slots(14).WhichNum).Enabled = False
            Else
                D(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 4 Then
                D(Slots(15).WhichNum).Enabled = False
            Else
                D(Slots(15).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(20).WhichNum).Left = 412 And S(Slots(20).WhichNum).Top = 200 Then
            If Slots(14).WhichCard = 1 Then
                H(Slots(14).WhichNum).Enabled = False
            Else
                H(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 1 Then
                H(Slots(15).WhichNum).Enabled = False
            Else
                H(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 2 Then
                S(Slots(14).WhichNum).Enabled = False
            Else
                S(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 2 Then
                S(Slots(15).WhichNum).Enabled = False
            Else
                S(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 3 Then
                C(Slots(14).WhichNum).Enabled = False
            Else
                C(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 3 Then
                C(Slots(15).WhichNum).Enabled = False
            Else
                C(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 4 Then
                D(Slots(14).WhichNum).Enabled = False
            Else
                D(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 4 Then
                D(Slots(15).WhichNum).Enabled = False
            Else
                D(Slots(15).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(20).WhichNum).Left = 412 And C(Slots(20).WhichNum).Top = 200 Then
            If Slots(14).WhichCard = 1 Then
                H(Slots(14).WhichNum).Enabled = False
            Else
                H(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 1 Then
                H(Slots(15).WhichNum).Enabled = False
            Else
                H(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 2 Then
                S(Slots(14).WhichNum).Enabled = False
            Else
                S(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 2 Then
                S(Slots(15).WhichNum).Enabled = False
            Else
                S(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 3 Then
                C(Slots(14).WhichNum).Enabled = False
            Else
                C(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 3 Then
                C(Slots(15).WhichNum).Enabled = False
            Else
                C(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 4 Then
                D(Slots(14).WhichNum).Enabled = False
            Else
                D(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 4 Then
                D(Slots(15).WhichNum).Enabled = False
            Else
                D(Slots(15).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(20).WhichNum).Left = 412 And D(Slots(20).WhichNum).Top = 200 Then
            If Slots(14).WhichCard = 1 Then
                H(Slots(14).WhichNum).Enabled = False
            Else
                H(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 1 Then
                H(Slots(15).WhichNum).Enabled = False
            Else
                H(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 2 Then
                S(Slots(14).WhichNum).Enabled = False
            Else
                S(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 2 Then
                S(Slots(15).WhichNum).Enabled = False
            Else
                S(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 3 Then
                C(Slots(14).WhichNum).Enabled = False
            Else
                C(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 3 Then
                C(Slots(15).WhichNum).Enabled = False
            Else
                C(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(14).WhichCard = 4 Then
                D(Slots(14).WhichNum).Enabled = False
            Else
                D(Slots(14).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 4 Then
                D(Slots(15).WhichNum).Enabled = False
            Else
                D(Slots(15).WhichNum).Enabled = True
            End If
        End If
        
        
        '21
        If H(Slots(21).WhichNum).Left = 484 And H(Slots(21).WhichNum).Top = 200 Then
            If Slots(15).WhichCard = 1 Then
                H(Slots(15).WhichNum).Enabled = False
            Else
                H(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 2 Then
                S(Slots(15).WhichNum).Enabled = False
            Else
                S(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 3 Then
                C(Slots(15).WhichNum).Enabled = False
            Else
                C(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 4 Then
                D(Slots(15).WhichNum).Enabled = False
            Else
                D(Slots(15).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(21).WhichNum).Left = 484 And S(Slots(21).WhichNum).Top = 200 Then
            If Slots(15).WhichCard = 1 Then
                H(Slots(15).WhichNum).Enabled = False
            Else
                H(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 2 Then
                S(Slots(15).WhichNum).Enabled = False
            Else
                S(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 3 Then
                C(Slots(15).WhichNum).Enabled = False
            Else
                C(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 4 Then
                D(Slots(15).WhichNum).Enabled = False
            Else
                D(Slots(15).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(21).WhichNum).Left = 484 And C(Slots(21).WhichNum).Top = 200 Then
            If Slots(15).WhichCard = 1 Then
                H(Slots(15).WhichNum).Enabled = False
            Else
                H(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 2 Then
                S(Slots(15).WhichNum).Enabled = False
            Else
                S(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 3 Then
                C(Slots(15).WhichNum).Enabled = False
            Else
                C(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 4 Then
                D(Slots(15).WhichNum).Enabled = False
            Else
                D(Slots(15).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(21).WhichNum).Left = 484 And D(Slots(21).WhichNum).Top = 200 Then
            If Slots(15).WhichCard = 1 Then
                H(Slots(15).WhichNum).Enabled = False
            Else
                H(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 2 Then
                S(Slots(15).WhichNum).Enabled = False
            Else
                S(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 3 Then
                C(Slots(15).WhichNum).Enabled = False
            Else
                C(Slots(15).WhichNum).Enabled = True
            End If
            If Slots(15).WhichCard = 4 Then
                D(Slots(15).WhichNum).Enabled = False
            Else
                D(Slots(15).WhichNum).Enabled = True
            End If
        End If
End Sub

Private Sub RowCheck7(intChecking As Integer)
'ROW 7------------------------------------
        '22
        If H(Slots(22).WhichNum).Left = 88 And H(Slots(22).WhichNum).Top = 225 Then
            If Slots(16).WhichCard = 1 Then
                H(Slots(16).WhichNum).Enabled = False
            Else
                H(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 2 Then
                S(Slots(16).WhichNum).Enabled = False
            Else
                S(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 3 Then
                C(Slots(16).WhichNum).Enabled = False
            Else
                C(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 4 Then
                D(Slots(16).WhichNum).Enabled = False
            Else
                D(Slots(16).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(22).WhichNum).Left = 88 And S(Slots(22).WhichNum).Top = 225 Then
            If Slots(16).WhichCard = 1 Then
                H(Slots(16).WhichNum).Enabled = False
            Else
                H(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 2 Then
                S(Slots(16).WhichNum).Enabled = False
            Else
                S(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 3 Then
                C(Slots(16).WhichNum).Enabled = False
            Else
                C(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 4 Then
                D(Slots(16).WhichNum).Enabled = False
            Else
                D(Slots(16).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(22).WhichNum).Left = 88 And C(Slots(22).WhichNum).Top = 225 Then
            If Slots(16).WhichCard = 1 Then
                H(Slots(16).WhichNum).Enabled = False
            Else
                H(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 2 Then
                S(Slots(16).WhichNum).Enabled = False
            Else
                S(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 3 Then
                C(Slots(16).WhichNum).Enabled = False
            Else
                C(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 4 Then
                D(Slots(16).WhichNum).Enabled = False
            Else
                D(Slots(16).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(22).WhichNum).Left = 88 And D(Slots(22).WhichNum).Top = 225 Then
            If Slots(16).WhichCard = 1 Then
                H(Slots(16).WhichNum).Enabled = False
            Else
                H(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 2 Then
                S(Slots(16).WhichNum).Enabled = False
            Else
                S(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 3 Then
                C(Slots(16).WhichNum).Enabled = False
            Else
                C(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 4 Then
                D(Slots(16).WhichNum).Enabled = False
            Else
                D(Slots(16).WhichNum).Enabled = True
            End If
        End If
        
        '23
        If D(Slots(23).WhichNum).Left = 160 And D(Slots(23).WhichNum).Top = 225 Then
            If Slots(16).WhichCard = 1 Then
                H(Slots(16).WhichNum).Enabled = False
            Else
                H(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 1 Then
                H(Slots(17).WhichNum).Enabled = False
            Else
                H(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 2 Then
                S(Slots(16).WhichNum).Enabled = False
            Else
                S(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 2 Then
                S(Slots(17).WhichNum).Enabled = False
            Else
                S(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 3 Then
                C(Slots(16).WhichNum).Enabled = False
            Else
                C(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 3 Then
                C(Slots(17).WhichNum).Enabled = False
            Else
                C(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 4 Then
                D(Slots(16).WhichNum).Enabled = False
            Else
                D(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 4 Then
                D(Slots(17).WhichNum).Enabled = False
            Else
                D(Slots(17).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(23).WhichNum).Left = 160 And C(Slots(23).WhichNum).Top = 225 Then
            If Slots(16).WhichCard = 1 Then
                H(Slots(16).WhichNum).Enabled = False
            Else
                H(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 1 Then
                H(Slots(17).WhichNum).Enabled = False
            Else
                H(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 2 Then
                S(Slots(16).WhichNum).Enabled = False
            Else
                S(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 2 Then
                S(Slots(17).WhichNum).Enabled = False
            Else
                S(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 3 Then
                C(Slots(16).WhichNum).Enabled = False
            Else
                C(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 3 Then
                C(Slots(17).WhichNum).Enabled = False
            Else
                C(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 4 Then
                D(Slots(16).WhichNum).Enabled = False
            Else
                D(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 4 Then
                D(Slots(17).WhichNum).Enabled = False
            Else
                D(Slots(17).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(23).WhichNum).Left = 160 And S(Slots(23).WhichNum).Top = 225 Then
            If Slots(16).WhichCard = 1 Then
                H(Slots(16).WhichNum).Enabled = False
            Else
                H(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 1 Then
                H(Slots(17).WhichNum).Enabled = False
            Else
                H(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 2 Then
                S(Slots(16).WhichNum).Enabled = False
            Else
                S(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 2 Then
                S(Slots(17).WhichNum).Enabled = False
            Else
                S(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 3 Then
                C(Slots(16).WhichNum).Enabled = False
            Else
                C(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 3 Then
                C(Slots(17).WhichNum).Enabled = False
            Else
                C(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 4 Then
                D(Slots(16).WhichNum).Enabled = False
            Else
                D(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 4 Then
                D(Slots(17).WhichNum).Enabled = False
            Else
                D(Slots(17).WhichNum).Enabled = True
            End If
        End If
        
        If H(Slots(23).WhichNum).Left = 160 And H(Slots(23).WhichNum).Top = 225 Then
            If Slots(16).WhichCard = 1 Then
                H(Slots(16).WhichNum).Enabled = False
            Else
                H(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 1 Then
                H(Slots(17).WhichNum).Enabled = False
            Else
                H(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 2 Then
                S(Slots(16).WhichNum).Enabled = False
            Else
                S(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 2 Then
                S(Slots(17).WhichNum).Enabled = False
            Else
                S(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 3 Then
                C(Slots(16).WhichNum).Enabled = False
            Else
                C(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 3 Then
                C(Slots(17).WhichNum).Enabled = False
            Else
                C(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(16).WhichCard = 4 Then
                D(Slots(16).WhichNum).Enabled = False
            Else
                D(Slots(16).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 4 Then
                D(Slots(17).WhichNum).Enabled = False
            Else
                D(Slots(17).WhichNum).Enabled = True
            End If
        End If
        
        '24
        If H(Slots(24).WhichNum).Left = 232 And H(Slots(24).WhichNum).Top = 225 Then
            If Slots(17).WhichCard = 1 Then
                H(Slots(17).WhichNum).Enabled = False
            Else
                H(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 1 Then
                H(Slots(18).WhichNum).Enabled = False
            Else
                H(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 2 Then
                S(Slots(17).WhichNum).Enabled = False
            Else
                S(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 2 Then
                S(Slots(18).WhichNum).Enabled = False
            Else
                S(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 3 Then
                C(Slots(17).WhichNum).Enabled = False
            Else
                C(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 3 Then
                C(Slots(18).WhichNum).Enabled = False
            Else
                C(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 4 Then
                D(Slots(17).WhichNum).Enabled = False
            Else
                D(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 4 Then
                D(Slots(18).WhichNum).Enabled = False
            Else
                D(Slots(18).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(24).WhichNum).Left = 232 And S(Slots(24).WhichNum).Top = 225 Then
            If Slots(17).WhichCard = 1 Then
                H(Slots(17).WhichNum).Enabled = False
            Else
                H(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 1 Then
                H(Slots(18).WhichNum).Enabled = False
            Else
                H(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 2 Then
                S(Slots(17).WhichNum).Enabled = False
            Else
                S(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 2 Then
                S(Slots(18).WhichNum).Enabled = False
            Else
                S(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 3 Then
                C(Slots(17).WhichNum).Enabled = False
            Else
                C(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 3 Then
                C(Slots(18).WhichNum).Enabled = False
            Else
                C(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 4 Then
                D(Slots(17).WhichNum).Enabled = False
            Else
                D(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 4 Then
                D(Slots(18).WhichNum).Enabled = False
            Else
                D(Slots(18).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(24).WhichNum).Left = 232 And C(Slots(24).WhichNum).Top = 225 Then
            If Slots(17).WhichCard = 1 Then
                H(Slots(17).WhichNum).Enabled = False
            Else
                H(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 1 Then
                H(Slots(18).WhichNum).Enabled = False
            Else
                H(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 2 Then
                S(Slots(17).WhichNum).Enabled = False
            Else
                S(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 2 Then
                S(Slots(18).WhichNum).Enabled = False
            Else
                S(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 3 Then
                C(Slots(17).WhichNum).Enabled = False
            Else
                C(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 3 Then
                C(Slots(18).WhichNum).Enabled = False
            Else
                C(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 4 Then
                D(Slots(17).WhichNum).Enabled = False
            Else
                D(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 4 Then
                D(Slots(18).WhichNum).Enabled = False
            Else
                D(Slots(18).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(24).WhichNum).Left = 232 And D(Slots(24).WhichNum).Top = 225 Then
            If Slots(17).WhichCard = 1 Then
                H(Slots(17).WhichNum).Enabled = False
            Else
                H(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 1 Then
                H(Slots(18).WhichNum).Enabled = False
            Else
                H(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 2 Then
                S(Slots(17).WhichNum).Enabled = False
            Else
                S(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 2 Then
                S(Slots(18).WhichNum).Enabled = False
            Else
                S(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 3 Then
                C(Slots(17).WhichNum).Enabled = False
            Else
                C(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 3 Then
                C(Slots(18).WhichNum).Enabled = False
            Else
                C(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(17).WhichCard = 4 Then
                D(Slots(17).WhichNum).Enabled = False
            Else
                D(Slots(17).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 4 Then
                D(Slots(18).WhichNum).Enabled = False
            Else
                D(Slots(18).WhichNum).Enabled = True
            End If
        End If
        
        
        '25
        If D(Slots(25).WhichNum).Left = 304 And D(Slots(25).WhichNum).Top = 225 Then
            If Slots(18).WhichCard = 1 Then
                H(Slots(18).WhichNum).Enabled = False
            Else
                H(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 1 Then
                H(Slots(19).WhichNum).Enabled = False
            Else
                H(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 2 Then
                S(Slots(18).WhichNum).Enabled = False
            Else
                S(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 2 Then
                S(Slots(19).WhichNum).Enabled = False
            Else
                S(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 3 Then
                C(Slots(18).WhichNum).Enabled = False
            Else
                C(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 3 Then
                C(Slots(19).WhichNum).Enabled = False
            Else
                C(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 4 Then
                D(Slots(18).WhichNum).Enabled = False
            Else
                D(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 4 Then
                D(Slots(19).WhichNum).Enabled = False
            Else
                D(Slots(19).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(25).WhichNum).Left = 304 And C(Slots(25).WhichNum).Top = 225 Then
            If Slots(18).WhichCard = 1 Then
                H(Slots(18).WhichNum).Enabled = False
            Else
                H(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 1 Then
                H(Slots(19).WhichNum).Enabled = False
            Else
                H(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 2 Then
                S(Slots(18).WhichNum).Enabled = False
            Else
                S(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 2 Then
                S(Slots(19).WhichNum).Enabled = False
            Else
                S(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 3 Then
                C(Slots(18).WhichNum).Enabled = False
            Else
                C(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 3 Then
                C(Slots(19).WhichNum).Enabled = False
            Else
                C(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 4 Then
                D(Slots(18).WhichNum).Enabled = False
            Else
                D(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 4 Then
                D(Slots(19).WhichNum).Enabled = False
            Else
                D(Slots(19).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(25).WhichNum).Left = 304 And S(Slots(25).WhichNum).Top = 225 Then
            If Slots(18).WhichCard = 1 Then
                H(Slots(18).WhichNum).Enabled = False
            Else
                H(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 1 Then
                H(Slots(19).WhichNum).Enabled = False
            Else
                H(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 2 Then
                S(Slots(18).WhichNum).Enabled = False
            Else
                S(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 2 Then
                S(Slots(19).WhichNum).Enabled = False
            Else
                S(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 3 Then
                C(Slots(18).WhichNum).Enabled = False
            Else
                C(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 3 Then
                C(Slots(19).WhichNum).Enabled = False
            Else
                C(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 4 Then
                D(Slots(18).WhichNum).Enabled = False
            Else
                D(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 4 Then
                D(Slots(19).WhichNum).Enabled = False
            Else
                D(Slots(19).WhichNum).Enabled = True
            End If
        End If
        
        If H(Slots(25).WhichNum).Left = 304 And H(Slots(25).WhichNum).Top = 225 Then
            If Slots(18).WhichCard = 1 Then
                H(Slots(18).WhichNum).Enabled = False
            Else
                H(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 1 Then
                H(Slots(19).WhichNum).Enabled = False
            Else
                H(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 2 Then
                S(Slots(18).WhichNum).Enabled = False
            Else
                S(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 2 Then
                S(Slots(19).WhichNum).Enabled = False
            Else
                S(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 3 Then
                C(Slots(18).WhichNum).Enabled = False
            Else
                C(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 3 Then
                C(Slots(19).WhichNum).Enabled = False
            Else
                C(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(18).WhichCard = 4 Then
                D(Slots(18).WhichNum).Enabled = False
            Else
                D(Slots(18).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 4 Then
                D(Slots(19).WhichNum).Enabled = False
            Else
                D(Slots(19).WhichNum).Enabled = True
            End If
        End If
        
        
        '26
        If D(Slots(26).WhichNum).Left = 376 And D(Slots(26).WhichNum).Top = 225 Then
            If Slots(19).WhichCard = 1 Then
                H(Slots(19).WhichNum).Enabled = False
            Else
                H(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 1 Then
                H(Slots(20).WhichNum).Enabled = False
            Else
                H(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 2 Then
                S(Slots(19).WhichNum).Enabled = False
            Else
                S(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 2 Then
                S(Slots(20).WhichNum).Enabled = False
            Else
                S(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 3 Then
                C(Slots(19).WhichNum).Enabled = False
            Else
                C(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 3 Then
                C(Slots(20).WhichNum).Enabled = False
            Else
                C(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 4 Then
                D(Slots(19).WhichNum).Enabled = False
            Else
                D(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 4 Then
                D(Slots(20).WhichNum).Enabled = False
            Else
                D(Slots(20).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(26).WhichNum).Left = 376 And C(Slots(26).WhichNum).Top = 225 Then
            If Slots(19).WhichCard = 1 Then
                H(Slots(19).WhichNum).Enabled = False
            Else
                H(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 1 Then
                H(Slots(20).WhichNum).Enabled = False
            Else
                H(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 2 Then
                S(Slots(19).WhichNum).Enabled = False
            Else
                S(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 2 Then
                S(Slots(20).WhichNum).Enabled = False
            Else
                S(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 3 Then
                C(Slots(19).WhichNum).Enabled = False
            Else
                C(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 3 Then
                C(Slots(20).WhichNum).Enabled = False
            Else
                C(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 4 Then
                D(Slots(19).WhichNum).Enabled = False
            Else
                D(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 4 Then
                D(Slots(20).WhichNum).Enabled = False
            Else
                D(Slots(20).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(26).WhichNum).Left = 376 And S(Slots(26).WhichNum).Top = 225 Then
            If Slots(19).WhichCard = 1 Then
                H(Slots(19).WhichNum).Enabled = False
            Else
                H(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 1 Then
                H(Slots(20).WhichNum).Enabled = False
            Else
                H(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 2 Then
                S(Slots(19).WhichNum).Enabled = False
            Else
                S(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 2 Then
                S(Slots(20).WhichNum).Enabled = False
            Else
                S(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 3 Then
                C(Slots(19).WhichNum).Enabled = False
            Else
                C(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 3 Then
                C(Slots(20).WhichNum).Enabled = False
            Else
                C(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 4 Then
                D(Slots(19).WhichNum).Enabled = False
            Else
                D(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 4 Then
                D(Slots(20).WhichNum).Enabled = False
            Else
                D(Slots(20).WhichNum).Enabled = True
            End If
        End If
        
        If H(Slots(26).WhichNum).Left = 376 And H(Slots(26).WhichNum).Top = 225 Then
            If Slots(19).WhichCard = 1 Then
                H(Slots(19).WhichNum).Enabled = False
            Else
                H(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 1 Then
                H(Slots(20).WhichNum).Enabled = False
            Else
                H(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 2 Then
                S(Slots(19).WhichNum).Enabled = False
            Else
                S(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 2 Then
                S(Slots(20).WhichNum).Enabled = False
            Else
                S(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 3 Then
                C(Slots(19).WhichNum).Enabled = False
            Else
                C(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 3 Then
                C(Slots(20).WhichNum).Enabled = False
            Else
                C(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(19).WhichCard = 4 Then
                D(Slots(19).WhichNum).Enabled = False
            Else
                D(Slots(19).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 4 Then
                D(Slots(20).WhichNum).Enabled = False
            Else
                D(Slots(20).WhichNum).Enabled = True
            End If
        End If
        
        
        '27
        If D(Slots(27).WhichNum).Left = 448 And D(Slots(27).WhichNum).Top = 225 Then
            If Slots(20).WhichCard = 1 Then
                H(Slots(20).WhichNum).Enabled = False
            Else
                H(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 1 Then
                H(Slots(21).WhichNum).Enabled = False
            Else
                H(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 2 Then
                S(Slots(20).WhichNum).Enabled = False
            Else
                S(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 2 Then
                S(Slots(21).WhichNum).Enabled = False
            Else
                S(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 3 Then
                C(Slots(20).WhichNum).Enabled = False
            Else
                C(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 3 Then
                C(Slots(21).WhichNum).Enabled = False
            Else
                C(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 4 Then
                D(Slots(20).WhichNum).Enabled = False
            Else
                D(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 4 Then
                D(Slots(21).WhichNum).Enabled = False
            Else
                D(Slots(21).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(27).WhichNum).Left = 448 And C(Slots(27).WhichNum).Top = 225 Then
            If Slots(20).WhichCard = 1 Then
                H(Slots(20).WhichNum).Enabled = False
            Else
                H(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 1 Then
                H(Slots(21).WhichNum).Enabled = False
            Else
                H(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 2 Then
                S(Slots(20).WhichNum).Enabled = False
            Else
                S(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 2 Then
                S(Slots(21).WhichNum).Enabled = False
            Else
                S(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 3 Then
                C(Slots(20).WhichNum).Enabled = False
            Else
                C(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 3 Then
                C(Slots(21).WhichNum).Enabled = False
            Else
                C(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 4 Then
                D(Slots(20).WhichNum).Enabled = False
            Else
                D(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 4 Then
                D(Slots(21).WhichNum).Enabled = False
            Else
                D(Slots(21).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(27).WhichNum).Left = 448 And S(Slots(27).WhichNum).Top = 225 Then
            If Slots(20).WhichCard = 1 Then
                H(Slots(20).WhichNum).Enabled = False
            Else
                H(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 1 Then
                H(Slots(21).WhichNum).Enabled = False
            Else
                H(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 2 Then
                S(Slots(20).WhichNum).Enabled = False
            Else
                S(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 2 Then
                S(Slots(21).WhichNum).Enabled = False
            Else
                S(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 3 Then
                C(Slots(20).WhichNum).Enabled = False
            Else
                C(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 3 Then
                C(Slots(21).WhichNum).Enabled = False
            Else
                C(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 4 Then
                D(Slots(20).WhichNum).Enabled = False
            Else
                D(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 4 Then
                D(Slots(21).WhichNum).Enabled = False
            Else
                D(Slots(21).WhichNum).Enabled = True
            End If
        End If
        
        If H(Slots(27).WhichNum).Left = 448 And H(Slots(27).WhichNum).Top = 225 Then
            If Slots(20).WhichCard = 1 Then
                H(Slots(20).WhichNum).Enabled = False
            Else
                H(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 1 Then
                H(Slots(21).WhichNum).Enabled = False
            Else
                H(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 2 Then
                S(Slots(20).WhichNum).Enabled = False
            Else
                S(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 2 Then
                S(Slots(21).WhichNum).Enabled = False
            Else
                S(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 3 Then
                C(Slots(20).WhichNum).Enabled = False
            Else
                C(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 3 Then
                C(Slots(21).WhichNum).Enabled = False
            Else
                C(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(20).WhichCard = 4 Then
                D(Slots(20).WhichNum).Enabled = False
            Else
                D(Slots(20).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 4 Then
                D(Slots(21).WhichNum).Enabled = False
            Else
                D(Slots(21).WhichNum).Enabled = True
            End If
        End If

        '28
        If H(Slots(28).WhichNum).Left = 520 And H(Slots(28).WhichNum).Top = 225 Then
            If Slots(21).WhichCard = 1 Then
                H(Slots(21).WhichNum).Enabled = False
            Else
                H(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 2 Then
                S(Slots(21).WhichNum).Enabled = False
            Else
                S(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 3 Then
                C(Slots(21).WhichNum).Enabled = False
            Else
                C(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 4 Then
                D(Slots(21).WhichNum).Enabled = False
            Else
                D(Slots(21).WhichNum).Enabled = True
            End If
        End If
        
        If S(Slots(28).WhichNum).Left = 520 And S(Slots(28).WhichNum).Top = 225 Then
            If Slots(21).WhichCard = 1 Then
                H(Slots(21).WhichNum).Enabled = False
            Else
                H(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 2 Then
                S(Slots(21).WhichNum).Enabled = False
            Else
                S(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 3 Then
                C(Slots(21).WhichNum).Enabled = False
            Else
                C(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 4 Then
                D(Slots(21).WhichNum).Enabled = False
            Else
                D(Slots(21).WhichNum).Enabled = True
            End If
        End If
        
        If C(Slots(28).WhichNum).Left = 520 And C(Slots(28).WhichNum).Top = 225 Then
            If Slots(21).WhichCard = 1 Then
                H(Slots(21).WhichNum).Enabled = False
            Else
                H(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 2 Then
                S(Slots(21).WhichNum).Enabled = False
            Else
                S(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 3 Then
                C(Slots(21).WhichNum).Enabled = False
            Else
                C(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 4 Then
                D(Slots(21).WhichNum).Enabled = False
            Else
                D(Slots(21).WhichNum).Enabled = True
            End If
        End If
        
        If D(Slots(28).WhichNum).Left = 520 And D(Slots(28).WhichNum).Top = 225 Then
            If Slots(21).WhichCard = 1 Then
                H(Slots(21).WhichNum).Enabled = False
            Else
                H(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 2 Then
                S(Slots(21).WhichNum).Enabled = False
            Else
                S(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 3 Then
                C(Slots(21).WhichNum).Enabled = False
            Else
                C(Slots(21).WhichNum).Enabled = True
            End If
            If Slots(21).WhichCard = 4 Then
                D(Slots(21).WhichNum).Enabled = False
            Else
                D(Slots(21).WhichNum).Enabled = True
            End If
        End If

End Sub

'this is the end of checking



Private Sub D_Click(Index As Integer)
    
    If blnFirstCard = True Then
        If blnSecondCard = False Then
            intSecondCard = Index
            blnSecondCard = True
            intSecondSuit = 2
            SC.Visible = True
            SC.ZOrder 0
            SC.Left = D(Index).Left + 30
            SC.Top = D(Index).Top + 50
        End If
    End If
    If blnFirstCard = False Then
        intFirstCard = Index
        blnFirstCard = True
        intFirstSuit = 2
        FC.Visible = True
        FC.ZOrder 0
        FC.Left = D(Index).Left + 30
        FC.Top = D(Index).Top + 50
    End If
    
End Sub

Private Sub D_DblClick(Index As Integer)
    intThirdSuit = 1
    intThirdCard = Index
    DeckPile.Enabled = True
    SC.Visible = False
    FC.Visible = False
    intFirstCard = 0
    intSecondCard = 0
    blnFirstCard = False
    blnSecondCard = False
End Sub

Private Sub DeckPile_Timer()
    If intThirdSuit = 1 Then
        If D(intThirdCard).Left < 300 And D(intThirdCard).Top = 350 Then
        If blnAnimate = True Then
            D(intThirdCard).Left = D(intThirdCard).Left + 1
            Call MakeAllNotAble
            'D(intThirdCard).ZOrder 0
            If D(intThirdCard).Left >= 300 Then Call MakeAllAble: DeckPile.Enabled = False
        Else
            D(intThirdCard).Left = 300
            MakeAllAble
            DeckPile.Enabled = False
        End If
        End If
    End If
    If intThirdSuit = 2 Then
        If C(intThirdCard).Left < 300 And C(intThirdCard).Top = 350 Then
        If blnAnimate = True Then
            C(intThirdCard).Left = C(intThirdCard).Left + 1
            Call MakeAllNotAble
            'C(intThirdCard).ZOrder 0
            If C(intThirdCard).Left >= 300 Then Call MakeAllAble: DeckPile.Enabled = False
        Else
            C(intThirdCard).Left = 300
            MakeAllAble
            DeckPile.Enabled = False
        End If
        End If
    End If
    If intThirdSuit = 3 Then
        If H(intThirdCard).Left < 300 And H(intThirdCard).Top = 350 Then
        If blnAnimate = True Then
            H(intThirdCard).Left = H(intThirdCard).Left + 1
            Call MakeAllNotAble
            'H(intThirdCard).ZOrder 0
            If H(intThirdCard).Left >= 300 Then Call MakeAllAble: DeckPile.Enabled = False
        Else
            H(intThirdCard).Left = 300
            MakeAllAble
            DeckPile.Enabled = False
        End If
        End If
    End If
    If intThirdSuit = 4 Then
        If S(intThirdCard).Left < 300 And S(intThirdCard).Top = 350 Then
        If blnAnimate = True Then
            S(intThirdCard).Left = S(intThirdCard).Left + 1
            Call MakeAllNotAble
            'S(intThirdCard).ZOrder 0
            If S(intThirdCard).Left >= 300 Then Call MakeAllAble: DeckPile.Enabled = False
        Else
            S(intThirdCard).Left = 300
            MakeAllAble
            DeckPile.Enabled = False
        End If
        End If
    End If
       
End Sub

Private Sub Form_Load()
Randomize
Row(1).StartLeft = (304 - 72)
Row(1).Top = 75
Row(1).EndLeft = 304
Row(2).StartLeft = (268 - 72)
Row(2).Top = 100
Row(2).EndLeft = 340
Row(3).StartLeft = (232 - 72)
Row(3).Top = 125
Row(3).EndLeft = 376
Row(4).StartLeft = (196 - 72)
Row(4).Top = 150
Row(4).EndLeft = 412
Row(5).StartLeft = (160 - 72)
Row(5).Top = 175
Row(5).EndLeft = 448
Row(6).StartLeft = (124 - 72)
Row(6).Top = 200
Row(6).EndLeft = 484
Row(7).StartLeft = (88 - 72)
Row(7).Top = 225
Row(7).EndLeft = 520

blnAnimate = True

intBGI = 1
intRow = 1
intB = 1
blnFirstCard = False
blnSecondCard = False
Shuffle
Slozi

For intA = 1 To 13
    D(intA).Picture = Diamond.ListImages.Item(intA).Picture
    H(intA).Picture = Heart.ListImages.Item(intA).Picture
    C(intA).Picture = Clover.ListImages.Item(intA).Picture
    S(intA).Picture = Spade.ListImages.Item(intA).Picture
Next intA

End Sub

Private Sub Slozi()


Do Until intRow > 7
    
        If CardNumber(intB) < 14 Then
        If intRow < 8 Then
            CardHeart(CardNumber(intB)) = CardNumber(intB)
            Row(intRow).StartLeft = Row(intRow).StartLeft + 72
            intD = intD + 1
            Slots(intD).Left = Row(intRow).StartLeft
            Slots(intD).Top = Row(intRow).Top
            Slots(intD).WhichCard = 1
            Slots(intD).WhichNum = CardNumber(intB)
            H(CardHeart(CardNumber(intB))).Left = Row(intRow).StartLeft
            H(CardHeart(CardNumber(intB))).Top = Row(intRow).Top
            H(CardHeart(CardNumber(intB))).ZOrder 0
            intB = intB + 1
            If Row(intRow).StartLeft >= Row(intRow).EndLeft Then intRow = intRow + 1
        End If
        End If
    
        If CardNumber(intB) > 13 And CardNumber(intB) < 27 Then
        If intRow < 8 Then
            CardSpade(CardNumber(intB) - 13) = CardNumber(intB) - 13
            Row(intRow).StartLeft = Row(intRow).StartLeft + 72
            intD = intD + 1
            Slots(intD).Left = Row(intRow).StartLeft
            Slots(intD).Top = Row(intRow).Top
            Slots(intD).WhichCard = 2
            Slots(intD).WhichNum = CardNumber(intB) - 13
            S(CardSpade(CardNumber(intB) - 13)).Left = Row(intRow).StartLeft
            S(CardSpade(CardNumber(intB) - 13)).Top = Row(intRow).Top
            S(CardSpade(CardNumber(intB) - 13)).ZOrder 0
            intB = intB + 1
            If Row(intRow).StartLeft >= Row(intRow).EndLeft Then intRow = intRow + 1
        End If
        End If
        
        If CardNumber(intB) > 26 And CardNumber(intB) < 40 Then
        If intRow < 8 Then
            CardClover(CardNumber(intB) - 26) = CardNumber(intB) - 26
            Row(intRow).StartLeft = Row(intRow).StartLeft + 72
            intD = intD + 1
            Slots(intD).Left = Row(intRow).StartLeft
            Slots(intD).Top = Row(intRow).Top
            Slots(intD).WhichCard = 3
            Slots(intD).WhichNum = CardNumber(intB) - 26
            C(CardClover(CardNumber(intB) - 26)).Left = Row(intRow).StartLeft
            C(CardClover(CardNumber(intB) - 26)).Top = Row(intRow).Top
            C(CardClover(CardNumber(intB) - 26)).ZOrder 0
            intB = intB + 1
            If Row(intRow).StartLeft >= Row(intRow).EndLeft Then intRow = intRow + 1
        End If
        End If
        
        If CardNumber(intB) > 39 And CardNumber(intB) <= 52 Then
        If intRow < 8 Then
            CardDimond(CardNumber(intB) - 39) = CardNumber(intB) - 39
            Row(intRow).StartLeft = Row(intRow).StartLeft + 72
            intD = intD + 1
            Slots(intD).Left = Row(intRow).StartLeft
            Slots(intD).Top = Row(intRow).Top
            Slots(intD).WhichCard = 4
            Slots(intD).WhichNum = CardNumber(intB) - 39
            D(CardDimond(CardNumber(intB) - 39)).Left = Row(intRow).StartLeft
            D(CardDimond(CardNumber(intB) - 39)).Top = Row(intRow).Top
            D(CardDimond(CardNumber(intB) - 39)).ZOrder 0
            intB = intB + 1
            If Row(intRow).StartLeft >= Row(intRow).EndLeft Then intRow = intRow + 1
        End If
        End If
        If intRow > 7 Then SloziRest
Loop

    
    

End Sub


Private Sub SloziRest()

For intB = 29 To 52

        If CardNumber(intB) < 14 Then
            CardHeart(CardNumber(intB)) = CardNumber(intB)
            H(CardHeart(CardNumber(intB))).Left = 200
            H(CardHeart(CardNumber(intB))).Top = 350
            H(CardHeart(CardNumber(intB))).ZOrder 0
        End If
    
        If CardNumber(intB) > 13 And CardNumber(intB) < 27 Then
            CardSpade(CardNumber(intB) - 13) = CardNumber(intB) - 13
            S(CardSpade(CardNumber(intB) - 13)).Left = 200
            S(CardSpade(CardNumber(intB) - 13)).Top = 350
            S(CardSpade(CardNumber(intB) - 13)).ZOrder 0
        End If
        
        If CardNumber(intB) > 26 And CardNumber(intB) < 40 Then
            CardClover(CardNumber(intB) - 26) = CardNumber(intB) - 26
            C(CardClover(CardNumber(intB) - 26)).Left = 200
            C(CardClover(CardNumber(intB) - 26)).Top = 350
            C(CardClover(CardNumber(intB) - 26)).ZOrder 0
        End If
        
        If CardNumber(intB) > 39 And CardNumber(intB) <= 52 Then
            CardDimond(CardNumber(intB) - 39) = CardNumber(intB) - 39
            D(CardDimond(CardNumber(intB) - 39)).Left = 200
            D(CardDimond(CardNumber(intB) - 39)).Top = 350
            D(CardDimond(CardNumber(intB) - 39)).ZOrder 0
        End If

Next intB

End Sub

Private Sub FTF_Timer()
If Frame1.Visible = True Then Frame1.ZOrder 0: Picture2.ZOrder 1
End Sub

Private Sub H_Click(Index As Integer)
    If blnFirstCard = True Then
        If blnSecondCard = False Then
            intSecondCard = Index
            blnSecondCard = True
            intSecondSuit = 4
            SC.Visible = True
            SC.ZOrder 0
            SC.Left = H(Index).Left + 30
            SC.Top = H(Index).Top + 50
        End If
    End If
    If blnFirstCard = False Then
        intFirstCard = Index
        blnFirstCard = True
        intFirstSuit = 4
        FC.Visible = True
        FC.ZOrder 0
        FC.Left = H(Index).Left + 30
        FC.Top = H(Index).Top + 50
    End If
End Sub

Private Sub H_DblClick(Index As Integer)
    intThirdSuit = 3
    intThirdCard = Index
    intFirstCard = 0
    intSecondCard = 0
    blnFirstCard = False
    blnSecondCard = False
    DeckPile.Enabled = True
    SC.Visible = False
    FC.Visible = False
End Sub

Private Sub Label5_Click()
Frame1.Visible = False
End Sub

Private Sub mnuAnimate_Click()
mnuAnimate.Checked = Not mnuAnimate.Checked
blnAnimate = mnuAnimate.Checked
End Sub

Private Sub mnuCBI_Click()
If intBGI = 1 Then intBGI = 2: Picture1.Picture = LoadPicture(App.Path & "\playing board3.bmp"): mnuCBI.Caption = "Change Background Image -- 2": Exit Sub
If intBGI = 2 Then intBGI = 1: Picture1.Picture = LoadPicture(App.Path & "\playing board4.bmp"): mnuCBI.Caption = "Change Background Image -- 1": Exit Sub
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuHelp_Click()
If Frame1.Visible = False Then Frame1.Visible = True
End Sub

Private Sub mnuNewGame_Click()
Randomize
Row(1).StartLeft = (304 - 72)
Row(1).Top = 75
Row(1).EndLeft = 304
Row(2).StartLeft = (268 - 72)
Row(2).Top = 100
Row(2).EndLeft = 340
Row(3).StartLeft = (232 - 72)
Row(3).Top = 125
Row(3).EndLeft = 376
Row(4).StartLeft = (196 - 72)
Row(4).Top = 150
Row(4).EndLeft = 412
Row(5).StartLeft = (160 - 72)
Row(5).Top = 175
Row(5).EndLeft = 448
Row(6).StartLeft = (124 - 72)
Row(6).Top = 200
Row(6).EndLeft = 484
Row(7).StartLeft = (88 - 72)
Row(7).Top = 225
Row(7).EndLeft = 520

intRow = 1
intB = 1
intD = 0
intTS = 0
intFirstCard = 0
intSecondCard = 0
blnFirstCard = False
blnSecondCard = False
Shuffle
Slozi
YouWin.Visible = False


End Sub

Private Sub MoveTimer_Timer()

    If intFirstSuit = 1 Then
    If blnAnimate = True Then
        If blnZorder1 = True Then C(intFirstCard).ZOrder 0
        blnZorder1 = False
        MakeAllNotAble
        If C(intFirstCard).Left <= 535 Then C(intFirstCard).Left = C(intFirstCard).Left + 1
        If C(intFirstCard).Top >= 22 Then C(intFirstCard).Top = C(intFirstCard).Top - 1
        If C(intFirstCard).Top <= 22 And C(intFirstCard).Left >= 535 Then
            intFirstCard = 0
            blnFirstCard = False
            MakeAllAble
            MoveTimer.Enabled = False
        End If
    Else
        C(intFirstCard).ZOrder 0
        C(intFirstCard).Left = 535
        C(intFirstCard).Top = 22
        intFirstCard = 0
        blnFirstCard = False
        MakeAllAble
        MoveTimer.Enabled = False
    End If
    End If
    
    If intFirstSuit = 2 Then
    If blnAnimate = True Then
        If blnZorder1 = True Then D(intFirstCard).ZOrder 0
        blnZorder1 = False
        MakeAllNotAble
        If D(intFirstCard).Left <= 535 Then D(intFirstCard).Left = D(intFirstCard).Left + 1
        If D(intFirstCard).Top >= 22 Then D(intFirstCard).Top = D(intFirstCard).Top - 1
        If D(intFirstCard).Top <= 22 And D(intFirstCard).Left >= 535 Then
            intFirstCard = 0
            blnFirstCard = False
            MakeAllAble
            MoveTimer.Enabled = False
        End If
    Else
        D(intFirstCard).ZOrder 0
        D(intFirstCard).Left = 535
        D(intFirstCard).Top = 22
        intFirstCard = 0
        blnFirstCard = False
        MakeAllAble
        MoveTimer.Enabled = False
    End If
    End If
    
    If intFirstSuit = 3 Then
    If blnAnimate = True Then
        If blnZorder1 = True Then S(intFirstCard).ZOrder 0
        blnZorder1 = False
        MakeAllNotAble
        If S(intFirstCard).Left <= 535 Then S(intFirstCard).Left = S(intFirstCard).Left + 1
        If S(intFirstCard).Top >= 22 Then S(intFirstCard).Top = S(intFirstCard).Top - 1
        If S(intFirstCard).Top <= 22 And S(intFirstCard).Left >= 535 Then
            intFirstCard = 0
            blnFirstCard = False
            MakeAllAble
            MoveTimer.Enabled = False
        End If
    Else
        S(intFirstCard).ZOrder 0
        S(intFirstCard).Left = 535
        S(intFirstCard).Top = 22
        intFirstCard = 0
        blnFirstCard = False
        MakeAllAble
        MoveTimer.Enabled = False
    End If
    End If

    If intFirstSuit = 4 Then
    If blnAnimate = True Then
        If blnZorder1 = True Then H(intFirstCard).ZOrder 0
        blnZorder1 = False
        MakeAllNotAble
        If H(intFirstCard).Left <= 535 Then H(intFirstCard).Left = H(intFirstCard).Left + 1
        If H(intFirstCard).Top >= 22 Then H(intFirstCard).Top = H(intFirstCard).Top - 1
        If H(intFirstCard).Top <= 22 And H(intFirstCard).Left >= 535 Then
            intFirstCard = 0
            blnFirstCard = False
            MakeAllAble
            MoveTimer.Enabled = False
        End If
    Else
        H(intFirstCard).ZOrder 0
        H(intFirstCard).Left = 535
        H(intFirstCard).Top = 22
        intFirstCard = 0
        blnFirstCard = False
        MakeAllAble
        MoveTimer.Enabled = False
    End If
    End If
    
End Sub

Private Sub MoveTimer1_Timer()

    If intSecondSuit = 1 Then
    If blnAnimate = True Then
        If blnZorder2 = True Then C(intSecondCard).ZOrder 0
        blnZorder2 = False
        MakeAllNotAble
        If C(intSecondCard).Left <= 535 Then C(intSecondCard).Left = C(intSecondCard).Left + 1
        If C(intSecondCard).Top >= 22 Then C(intSecondCard).Top = C(intSecondCard).Top - 1
        If C(intSecondCard).Top <= 22 And C(intSecondCard).Left >= 535 Then
            intSecondCard = 0
            blnSecondCard = False
            MakeAllAble
            MoveTimer1.Enabled = False
        End If
    Else
        C(intSecondCard).ZOrder 0
        C(intSecondCard).Left = 535
        C(intSecondCard).Top = 22
        intSecondCard = 0
        blnSecondCard = False
        MakeAllAble
        MoveTimer1.Enabled = False
    End If
    End If
    
    If intSecondSuit = 2 Then
    If blnAnimate = True Then
        If blnZorder2 = True Then D(intSecondCard).ZOrder 0
        blnZorder2 = False
        MakeAllNotAble
        If D(intSecondCard).Left <= 535 Then D(intSecondCard).Left = D(intSecondCard).Left + 1
        If D(intSecondCard).Top >= 22 Then D(intSecondCard).Top = D(intSecondCard).Top - 1
        If D(intSecondCard).Top <= 22 And D(intSecondCard).Left >= 535 Then
            intSecondCard = 0
            blnSecondCard = False
            MakeAllAble
            MoveTimer1.Enabled = False
        End If
    Else
        D(intSecondCard).ZOrder 0
        D(intSecondCard).Left = 535
        D(intSecondCard).Top = 22
        intSecondCard = 0
        blnSecondCard = False
        MakeAllAble
        MoveTimer1.Enabled = False
    End If
    End If
    
    If intSecondSuit = 3 Then
    If blnAnimate = True Then
        If blnZorder2 = True Then S(intSecondCard).ZOrder 0
        blnZorder2 = False
        MakeAllNotAble
        If S(intSecondCard).Left <= 535 Then S(intSecondCard).Left = S(intSecondCard).Left + 1
        If S(intSecondCard).Top >= 22 Then S(intSecondCard).Top = S(intSecondCard).Top - 1
        If S(intSecondCard).Top <= 22 And S(intSecondCard).Left >= 535 Then
            intSecondCard = 0
            blnSecondCard = False
            MakeAllAble
            MoveTimer1.Enabled = False
        End If
    Else
        S(intSecondCard).ZOrder 0
        S(intSecondCard).Left = 535
        S(intSecondCard).Top = 22
        intSecondCard = 0
        blnSecondCard = False
        MakeAllAble
        MoveTimer1.Enabled = False
    End If
    End If
    
    If intSecondSuit = 4 Then
    If blnAnimate = True Then
        If blnZorder2 = True Then H(intSecondCard).ZOrder 0
        blnZorder2 = False
        MakeAllNotAble
        If H(intSecondCard).Left <= 535 Then H(intSecondCard).Left = H(intSecondCard).Left + 1
        If H(intSecondCard).Top >= 22 Then H(intSecondCard).Top = H(intSecondCard).Top - 1
        If H(intSecondCard).Top <= 22 And H(intSecondCard).Left >= 535 Then
            intSecondCard = 0
            blnSecondCard = False
            MakeAllAble
            MoveTimer1.Enabled = False
        End If
    Else
        H(intSecondCard).ZOrder 0
        H(intSecondCard).Left = 535
        H(intSecondCard).Top = 22
        intSecondCard = 0
        blnSecondCard = False
        MakeAllAble
        MoveTimer1.Enabled = False
    End If
    End If
    
End Sub

Private Sub ReDeck_Click()

For intB = 29 To 52
        
        If CardNumber(intB) < 14 Then
            CardHeart(CardNumber(intB)) = CardNumber(intB)
            If H(CardHeart(CardNumber(intB))).Top = 350 Then
                H(CardHeart(CardNumber(intB))).Left = 200
                H(CardHeart(CardNumber(intB))).Top = 350
                H(CardHeart(CardNumber(intB))).ZOrder 0
            End If
        End If
    
        If CardNumber(intB) > 13 And CardNumber(intB) < 27 Then
            CardSpade(CardNumber(intB) - 13) = CardNumber(intB) - 13
            If S(CardSpade(CardNumber(intB) - 13)).Top = 350 Then
                S(CardSpade(CardNumber(intB) - 13)).Left = 200
                S(CardSpade(CardNumber(intB) - 13)).Top = 350
                S(CardSpade(CardNumber(intB) - 13)).ZOrder 0
            End If
        End If
        
        If CardNumber(intB) > 26 And CardNumber(intB) < 40 Then
            CardClover(CardNumber(intB) - 26) = CardNumber(intB) - 26
            If C(CardHeart(CardNumber(intB) - 26)).Top = 350 Then
                C(CardClover(CardNumber(intB) - 26)).Left = 200
                C(CardClover(CardNumber(intB) - 26)).Top = 350
                C(CardClover(CardNumber(intB) - 26)).ZOrder 0
            End If
        End If
        
        If CardNumber(intB) > 39 And CardNumber(intB) <= 52 Then
            CardDimond(CardNumber(intB) - 39) = CardNumber(intB) - 39
            If D(CardHeart(CardNumber(intB) - 39)).Top = 350 Then
                D(CardDimond(CardNumber(intB) - 39)).Left = 200
                D(CardDimond(CardNumber(intB) - 39)).Top = 350
                D(CardDimond(CardNumber(intB) - 39)).ZOrder 0
            End If
        End If

Next intB

End Sub

Private Sub S_Click(Index As Integer)
    If blnFirstCard = True Then
        If blnSecondCard = False Then
            intSecondCard = Index
            blnSecondCard = True
            intSecondSuit = 3
            SC.Visible = True
            SC.ZOrder 0
            SC.Left = S(Index).Left + 30
            SC.Top = S(Index).Top + 50
        End If
    End If
    If blnFirstCard = False Then
        intFirstCard = Index
        blnFirstCard = True
        intFirstSuit = 3
        FC.Visible = True
        FC.ZOrder 0
        FC.Left = S(Index).Left + 30
        FC.Top = S(Index).Top + 50
    End If
End Sub

Private Sub S_DblClick(Index As Integer)
    intThirdSuit = 4
    intThirdCard = Index
    DeckPile.Enabled = True
    SC.Visible = False
    FC.Visible = False
    intFirstCard = 0
    intSecondCard = 0
    blnFirstCard = False
    blnSecondCard = False
End Sub

Private Sub Score_Timer()
intScore = intFirstCard + intSecondCard
    
    If blnFirstCard = True And intScore = 13 Then
        MoveTimer.Enabled = True
        If SC.Visible = True Then intTS = intTS + 13
        blnZorder1 = True
        SC.Visible = False
        FC.Visible = False
    End If

    If blnFirstCard = True And blnSecondCard = True And intScore <> 13 Then
        blnFirstCard = False
        blnSecondCard = False
        intFirstCard = 0
        intSecondCard = 0
        SC.Visible = False
        FC.Visible = False
    End If
    
    If blnFirstCard = True And blnSecondCard = True And intScore = 13 Then
        MoveTimer.Enabled = True
        MoveTimer1.Enabled = True
        If SC.Visible = True Then intTS = intTS + 13
        blnZorder1 = True
        blnZorder2 = True
        SC.Visible = False
        FC.Visible = False
    End If
    
    Label2.Caption = "Score: " & intTS
    Label1.Caption = intScore
End Sub

Private Sub MakeAllNotAble()
    For intC = 1 To 13

        D(intC).Enabled = False
        H(intC).Enabled = False
        S(intC).Enabled = False
        C(intC).Enabled = False
        
    Next intC
End Sub

Private Sub MakeAllAble()
    For intC = 1 To 13

        D(intC).Enabled = True
        H(intC).Enabled = True
        S(intC).Enabled = True
        C(intC).Enabled = True
        
    Next intC
End Sub






Private Sub Win_Timer()
If Slots(1).WhichCard = 1 Then
    If H(Slots(1).WhichNum).Left >= 530 And H(Slots(1).WhichNum).Top <= 25 Then
        YouWin.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        YouWin.Visible = True
    End If
End If
    
If Slots(1).WhichCard = 2 Then
    If S(Slots(1).WhichNum).Left >= 530 And S(Slots(1).WhichNum).Top <= 25 Then
        YouWin.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        YouWin.Visible = True
    End If
End If

If Slots(1).WhichCard = 3 Then
    If C(Slots(1).WhichNum).Left >= 530 And C(Slots(1).WhichNum).Top <= 25 Then
        YouWin.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        YouWin.Visible = True
    End If
End If

If Slots(1).WhichCard = 4 Then
    If D(Slots(1).WhichNum).Left >= 530 And D(Slots(1).WhichNum).Top <= 25 Then
        YouWin.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        YouWin.Visible = True
    End If
End If
    
End Sub
