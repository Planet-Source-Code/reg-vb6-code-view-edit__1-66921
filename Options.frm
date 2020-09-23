VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00685758&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Editor Options"
   ClientHeight    =   8415
   ClientLeft      =   4065
   ClientTop       =   1980
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Euclid"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   5190
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4020
      TabIndex        =   36
      Top             =   45
      Width           =   1005
   End
   Begin VB.Frame Frame4 
      Caption         =   "Text Colours"
      Height          =   4875
      Left            =   180
      TabIndex        =   13
      Top             =   3300
      Width           =   4875
      Begin VB.Label LbCol 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "For Each Obj"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   7
         Left            =   2910
         TabIndex        =   41
         Top             =   2610
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Edited Lines"
         Height          =   345
         Index           =   7
         Left            =   60
         TabIndex        =   40
         Top             =   2550
         Width           =   2325
      End
      Begin VB.Label LbColPal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aa"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Index           =   7
         Left            =   3990
         TabIndex        =   39
         Top             =   3990
         Width           =   525
      End
      Begin VB.Label LbColPal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aa"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   6
         Left            =   3465
         TabIndex        =   38
         Top             =   3990
         Width           =   525
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click on the item you wish to change the color of in the list above. Then select the color you want from the strip below "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   285
         TabIndex        =   34
         Top             =   3030
         Width           =   4305
      End
      Begin VB.Label LbColPal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aa"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   5
         Left            =   2925
         TabIndex        =   33
         Top             =   3990
         Width           =   525
      End
      Begin VB.Label LbColPal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aa"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   4
         Left            =   2400
         TabIndex        =   32
         Top             =   3990
         Width           =   525
      End
      Begin VB.Label LbColPal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aa"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   3
         Left            =   1860
         TabIndex        =   31
         Top             =   3990
         Width           =   525
      End
      Begin VB.Label LbColPal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aa"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   2
         Left            =   1335
         TabIndex        =   30
         Top             =   3990
         Width           =   525
      End
      Begin VB.Label LbColPal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aa"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   795
         TabIndex        =   29
         Top             =   3990
         Width           =   525
      End
      Begin VB.Label LbColPal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aa"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   28
         Top             =   3990
         Width           =   525
      End
      Begin VB.Label LbCol 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "For Each Obj"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   6
         Left            =   2910
         TabIndex        =   27
         Top             =   2310
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Height          =   345
         Index           =   6
         Left            =   60
         TabIndex        =   26
         Top             =   2250
         Width           =   2325
      End
      Begin VB.Label LbCol 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "nSum = nTot"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   5
         Left            =   2910
         TabIndex        =   25
         Top             =   1995
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Other Code Lines"
         Height          =   345
         Index           =   5
         Left            =   60
         TabIndex        =   24
         Top             =   1935
         Width           =   2325
      End
      Begin VB.Label LbCol 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dim strName"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   4
         Left            =   2910
         TabIndex        =   23
         Top             =   1635
         Width           =   1695
      End
      Begin VB.Label LbCol 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Public gnCount"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   3
         Left            =   2910
         TabIndex        =   22
         Top             =   1290
         Width           =   1695
      End
      Begin VB.Label LbCol 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "' Sets all colors"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Index           =   2
         Left            =   2910
         TabIndex        =   21
         Top             =   975
         Width           =   1695
      End
      Begin VB.Label LbCol 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Form Load"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   2910
         TabIndex        =   20
         Top             =   675
         Width           =   1695
      End
      Begin VB.Label LbCol 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "frmVBEditor"
         BeginProperty Font 
            Name            =   "Euclid"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   2910
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Module Variables"
         Height          =   345
         Index           =   4
         Left            =   60
         TabIndex        =   18
         Top             =   1605
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Public Variables"
         Height          =   345
         Index           =   3
         Left            =   60
         TabIndex        =   17
         Top             =   1260
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks"
         Height          =   345
         Index           =   2
         Left            =   60
         TabIndex        =   16
         Top             =   945
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Procedure Names"
         Height          =   345
         Index           =   1
         Left            =   60
         TabIndex        =   15
         Top             =   645
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Module Names"
         Height          =   345
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   360
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bold"
      Height          =   1845
      Left            =   3930
      TabIndex        =   10
      Top             =   450
      Width           =   1095
      Begin VB.OptionButton OptB 
         Caption         =   "No"
         CausesValidation=   0   'False
         Height          =   390
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1140
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton OptB 
         Caption         =   "Yes"
         CausesValidation=   0   'False
         Height          =   390
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   915
      End
   End
   Begin VB.TextBox txtExample 
      Alignment       =   2  'Center
      Height          =   510
      Left            =   1200
      TabIndex        =   9
      Text            =   "AbCd ""0123"" ~ # & =*(strW$) %"
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Size"
      Height          =   1845
      Left            =   2820
      TabIndex        =   4
      Top             =   450
      Width           =   1125
      Begin VB.OptionButton OptS 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   270
         TabIndex        =   35
         Top             =   1470
         Width           =   1095
      End
      Begin VB.OptionButton OptS 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   270
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton OptS 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   270
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptS 
         Caption         =   " 8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Top             =   390
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Font Name"
      Height          =   1845
      Left            =   180
      TabIndex        =   0
      Top             =   450
      Width           =   2685
      Begin VB.OptionButton Opt 
         Caption         =   "Courier New"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   1320
         Width           =   2205
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Times New Roman"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   900
         Width           =   2625
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Euclid"
         Height          =   390
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   450
         Value           =   -1  'True
         Width           =   2205
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Changes are Automatically Saved"
      BeginProperty Font 
         Name            =   "Euclid"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   180
      TabIndex        =   37
      Top             =   30
      Width           =   3795
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Example"
      Height          =   735
      Left            =   150
      TabIndex        =   8
      Top             =   2430
      Width           =   4875
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim nColInd             As Long 'Index of the Item to change
Private Sub cmdExit_Click()
    Unload Me
    End Sub
Private Sub Form_Load()
    SetOptions
    ShowExample
    End Sub
Private Sub SetOptions()
    ' The pupose of this Sub is to make the Options
    ' Buttons match the values held by the variables
    Dim N       As Long
    'FontName
    For N = 0 To 2
        If Opt(N).Caption = sfName Then
            Opt(N).Value = True
        End If
    Next
    'Font Size
    For N = 0 To 3
        If Val(OptS(N).Caption) = nFSize Then
            OptS(N).Value = True
        End If
    Next
    'Bold or NOT Bold
    If bBold = True Then
        OptB(0).Value = True
    Else
        OptB(1).Value = True
    End If
    'The Forecolors for various code lines
    For N = 0 To 6
        LbCol(N).ForeColor = gnCol(N)
    Next
    End Sub
Private Sub LbCol_Click(Index As Integer)
    ' Variable holds the index of the selected item
    nColInd = Index
    End Sub
Private Sub LbColPal_Click(Index As Integer)
    ' Set global variable and show effect on label
    gnCol(nColInd) = LbColPal(Index).ForeColor
    LbCol(nColInd).ForeColor = gnCol(nColInd)
    bColors = True
    End Sub
Private Sub Opt_Click(Index As Integer)
    ' Select FontName and show effect
    sfName = Opt(Index).Caption
    ShowExample
    End Sub
Private Sub OptB_Click(Index As Integer)
    ' Select Bold Yes/No and show effect
    If Index = 0 Then
        bBold = True
    Else
        bBold = False
    End If
    ShowExample
    End Sub
Private Sub OptS_Click(Index As Integer)
    ' Select Font Size and show effect
    nFSize = Val(OptS(Index).Caption)
    ShowExample
    End Sub
Private Sub ShowExample()
    ' Show results of selections
    txtExample.FontName = sfName
    txtExample.FontSize = nFSize
    txtExample.FontBold = bBold
    End Sub
Private Sub Picture1_Click()
    End Sub
