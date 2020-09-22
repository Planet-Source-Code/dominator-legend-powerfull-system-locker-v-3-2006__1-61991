VERSION 5.00
Begin VB.Form Frm_Main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5715
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15345
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   ForeColor       =   &H00808080&
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   15345
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic_Interface 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5715
      Index           =   0
      Left            =   0
      ScaleHeight     =   5715
      ScaleWidth      =   15345
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15345
      Begin VB.PictureBox Pic_Interface 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   2
         Left            =   0
         Picture         =   "Frm_Main.frx":628A
         ScaleHeight     =   390
         ScaleWidth      =   43425
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   43425
      End
      Begin VB.PictureBox Pic_Interface 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   5
         Left            =   0
         Picture         =   "Frm_Main.frx":3D52C
         ScaleHeight     =   390
         ScaleWidth      =   43425
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   5340
         Width           =   43425
         Begin VB.Label Lbl_Interface 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "© 2005, Dominator Legend"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   90
            Width           =   1995
         End
      End
      Begin VB.PictureBox Pic_Interface 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H003E52FD&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   3
         Left            =   0
         Picture         =   "Frm_Main.frx":747CE
         ScaleHeight     =   795
         ScaleWidth      =   15360
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   390
         Width           =   15360
      End
      Begin VB.PictureBox Pic_Interface 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   6
         Left            =   0
         Picture         =   "Frm_Main.frx":9C412
         ScaleHeight     =   390
         ScaleWidth      =   43425
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1170
         Width           =   43425
      End
      Begin VB.Timer Tim_Show 
         Interval        =   1
         Left            =   3570
         Top             =   1530
      End
      Begin VB.TextBox Txt_Interface 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1215
         Left            =   4320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "Frm_Main.frx":D36B4
         Top             =   1680
         Width           =   10785
      End
      Begin VB.PictureBox Pic_Interface 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3825
         Index           =   1
         Left            =   -120
         Picture         =   "Frm_Main.frx":D3821
         ScaleHeight     =   3825
         ScaleWidth      =   4245
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4245
         Begin VB.Timer Detector 
            Enabled         =   0   'False
            Interval        =   2000
            Left            =   3690
            Top             =   390
         End
      End
      Begin VB.PictureBox Pic_Con1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2205
         Left            =   4320
         ScaleHeight     =   2205
         ScaleWidth      =   10785
         TabIndex        =   8
         Top             =   3135
         Width           =   10785
         Begin VB.PictureBox Pic_Interface 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   1890
            Index           =   4
            Left            =   4530
            ScaleHeight     =   1890
            ScaleWidth      =   5145
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   330
            Width           =   5145
            Begin VB.PictureBox Pic_Error 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   1485
               Left            =   0
               ScaleHeight     =   1455
               ScaleWidth      =   5055
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   5085
               Begin VB.PictureBox Pic_Interface 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   375
                  Index           =   7
                  Left            =   -30
                  Picture         =   "Frm_Main.frx":108911
                  ScaleHeight     =   375
                  ScaleWidth      =   5115
                  TabIndex        =   14
                  TabStop         =   0   'False
                  Top             =   -30
                  Width           =   5115
                  Begin VB.Label Lbl_Interface 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Server Messages"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C0C0FF&
                     Height          =   195
                     Index           =   2
                     Left            =   135
                     TabIndex        =   15
                     Top             =   60
                     Width           =   1245
                  End
               End
               Begin VB.Image Img_Interface 
                  Height          =   690
                  Left            =   240
                  Picture         =   "Frm_Main.frx":11D9B3
                  Top             =   570
                  Width           =   600
               End
               Begin VB.Label Lbl_Message 
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H006C57FF&
                  Height          =   765
                  Left            =   1050
                  TabIndex        =   16
                  Top             =   570
                  Width           =   3810
               End
            End
         End
         Begin VB.CommandButton Cmd_Execute 
            Caption         =   "Unlock Windows"
            Height          =   375
            Left            =   2610
            TabIndex        =   11
            Top             =   1440
            Width           =   1755
         End
         Begin VB.TextBox Txt_Username 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   1680
            TabIndex        =   10
            Text            =   "Dominator"
            Top             =   330
            Width           =   2685
         End
         Begin VB.TextBox Txt_Password 
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   6
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "="
            TabIndex        =   9
            Top             =   870
            Width           =   2685
         End
         Begin VB.Label Lbl_Interface 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   18
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Lbl_Interface 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Password:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   17
            Top             =   930
            Width           =   1125
         End
      End
      Begin VB.PictureBox Pic_Authorized 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2205
         Left            =   4320
         ScaleHeight     =   2205
         ScaleWidth      =   10785
         TabIndex        =   19
         Top             =   3135
         Visible         =   0   'False
         Width           =   10785
         Begin VB.TextBox Txt_Auz1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   345
            Index           =   0
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "Frm_Main.frx":11EF85
            Top             =   720
            Width           =   10785
         End
         Begin VB.TextBox Txt_Auz1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   1
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "Frm_Main.frx":11EF9D
            Top             =   1080
            Width           =   10785
         End
      End
      Begin VB.Line Lin_Interface 
         BorderColor     =   &H00C0C0FF&
         X1              =   4320
         X2              =   15120
         Y1              =   3060
         Y2              =   3060
      End
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem -> ***|*********************************************************************|***|
Rem -> ***|                                                                     |***|
Rem -> ***|                      ______                __                       |***|
Rem -> ***|                     | ____ \              |  |                      |***|
Rem -> ***|                     | |   \ \             |  |                      |***|
Rem -> ***|                     | |    \ \            |  |                      |***|
Rem -> ***|                     | |    / /            |  |                      |***|
Rem -> ***|                     | |___/ /     __      |  |______                |***|
Rem -> ***|                     |______/     (__)     |_________|               |***|
Rem -> ***|                                                                     |***|
Rem -> ***|   _______________________________________________________________   |***|
Rem -> ***|                                                                     |***|
Rem -> ***|   Author       : John Fawzy (Dominator Legend)                      |***|
Rem -> ***|   Email        : Dominator_Legand@Yahoo.com                         |***|
Rem -> ***|   Date         : 21/3/2006                                          |***|
Rem -> ***|   Copyrights   : Some of these function not written by me,          |***|
Rem -> ***|                  However, Contents of code must be intact without   |***|
Rem -> ***|                  Change, If this work will used for commercial      |***|
Rem -> ***|                  Purpose please inform me, if you like this code    |***|
Rem -> ***|                  Please Rate It, Thanks                             |***|
Rem -> ***|                                                                     |***|
Rem -> ***|*********************************************************************|***|
Option Explicit
Dim ShowError       As Boolean
Private Sub Detector_Timer()
    Call ShellController(False)
End Sub
Private Sub Form_Initialize()
    If App.PrevInstance Then Beep: End
    Call MinimizeAllWindows: Working (1000)
    Pic_Error.Top = 1890
End Sub
Private Sub Form_Load()
    Me.Top = (Screen.Height / 2) - (Me.ScaleHeight / 2): Me.Left = (Screen.Width / 2) - (Me.ScaleWidth / 2)
    Me.Width = Pic_Interface(0).Width: Me.Height = Pic_Interface(0).Height
    Me.Show: DoEvents
    Call EnumWindows(AddressOf WndProc, 0): Working (100)
    Call EnableWindows(False, Me.HWnd)
    Call SetWindowPos(Me.HWnd, -1, 0, 0, 0, 0, &H2 + &H1)
    Call SetHook
    Detector.Enabled = True
    Txt_Username.SetFocus
End Sub
Private Sub Cmd_Execute_Click()
    If Txt_Username.Text = "" Then Txt_Username.SetFocus: InitError "Username Required.": Exit Sub
    If Txt_Password.Text = "" Then Txt_Password.SetFocus: InitError "Password Required.": Exit Sub
    If Txt_Username.Text = "Dominator" And Txt_Password.Text = "0" Then
        Pic_Con1.Visible = False
        Pic_Authorized.Visible = True
        Rem-> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        WindowsEnabled = True
        Detector.Enabled = False
        Call EnumWindows(AddressOf WndProc, 0)
        Call EnableWindows(True, Me.HWnd)
        Call UnSetHook
        Rem-> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Working (5000)
        TerminateProcess GetCurrentProcess, 0
        End
    Else
        InitError "Server report that either user name or password is invalide, Please try again."
    End If
End Sub
Private Sub Tim_Show_Timer()
    If ShowError Then
        If Pic_Error.Top <= 0 Then
            Pic_Error.Top = 0
            Tim_Show.Enabled = False: Tim_Show.Interval = 6000: Tim_Show.Enabled = True
            ShowError = False
            Exit Sub
        End If
        Pic_Error.Top = Pic_Error.Top - 50
    Else
        If Pic_Error.Top = 0 Then
            Tim_Show.Enabled = False: Tim_Show.Interval = 1: Tim_Show.Enabled = True
        End If
        If Pic_Error.Top >= 1890 Then
            Pic_Error.Top = 1890
            Tim_Show.Enabled = False
            ShowError = True
        End If
        Pic_Error.Top = Pic_Error.Top + 50
    End If
End Sub
Private Sub Txt_Password_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Cmd_Execute = True
End Sub
Private Sub Txt_Username_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Cmd_Execute = True
End Sub
Private Sub InitError(ErrorMessage As String)
    Beep
    Lbl_Message.Caption = ErrorMessage
    ShowError = True
    Pic_Error.Visible = True
    Pic_Error.Top = 1890
    Tim_Show.Enabled = False: Tim_Show.Interval = 1: Tim_Show.Enabled = True
End Sub
