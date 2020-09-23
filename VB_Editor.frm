VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H0050BEBE&
   Caption         =   "VB Editor   - by PointEqual"
   ClientHeight    =   6555
   ClientLeft      =   2625
   ClientTop       =   6585
   ClientWidth     =   7830
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7830
   Begin MSComctlLib.TreeView TRV 
      Height          =   405
      Left            =   -120
      TabIndex        =   0
      Top             =   900
      Visible         =   0   'False
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   714
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "IMG"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Euclid"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6090
      TabIndex        =   9
      Top             =   7230
      Width           =   1155
   End
   Begin VB.Frame fraCover 
      BackColor       =   &H00685758&
      BorderStyle     =   0  'None
      Caption         =   "       Module Selection       "
      ForeColor       =   &H0080FFFF&
      Height          =   885
      Left            =   30
      TabIndex        =   14
      Top             =   30
      Width           =   9915
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H009AD6E2&
         BackStyle       =   0  'Transparent
         Caption         =   "that you wish to work with then"
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   5580
         TabIndex        =   17
         Top             =   60
         Width           =   3075
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H009AD6E2&
         BackStyle       =   0  'Transparent
         Caption         =   "Select the Drive and Path for the VB Application"
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   570
         TabIndex        =   16
         Top             =   60
         Width           =   4875
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H009AD6E2&
         BackStyle       =   0  'Transparent
         Caption         =   "Check the Module(s) that you wish to View/Edit"
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   2100
         TabIndex        =   15
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      ToolTipText     =   "Save this module with the changes made"
      Top             =   30
      Width           =   1155
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert  >"
      Height          =   375
      Index           =   1
      Left            =   1230
      TabIndex        =   12
      ToolTipText     =   "Insert an INDENTED line below the selected line "
      Top             =   30
      Width           =   1155
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      ToolTipText     =   "Delete the selected line "
      Top             =   30
      Width           =   1155
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert "
      Height          =   375
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Insert a line below the selected line "
      Top             =   30
      Width           =   1155
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   375
      Left            =   4860
      TabIndex        =   8
      Top             =   7230
      Width           =   1155
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      ToolTipText     =   "Change FointName Size and Colors"
      Top             =   30
      Width           =   1155
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   10560
      Top             =   1410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VB_Editor.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VB_Editor.frx":0772
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VB_Editor.frx":0BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VB_Editor.frx":101E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VB_Editor.frx":1470
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VB_Editor.frx":18C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VB_Editor.frx":1D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VB_Editor.frx":2166
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VB_Editor.frx":25B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkSort 
      Alignment       =   1  'Right Justify
      Caption         =   "Sort Procedures"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4860
      TabIndex        =   18
      Top             =   6690
      Value           =   1  'Checked
      Width           =   2385
   End
   Begin VB.CommandButton cmdDirty 
      BackColor       =   &H80000005&
      Height          =   495
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   435
      Width           =   165
   End
   Begin VB.TextBox txtEdit 
      BeginProperty Font 
         Name            =   "Euclid"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   270
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "VB_Editor.frx":2D2A
      Top             =   420
      Width           =   8955
   End
   Begin VB.DriveListBox DRV 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1740
      TabIndex        =   1
      Top             =   1950
      Width           =   2955
   End
   Begin VB.DirListBox DIR 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   1740
      TabIndex        =   2
      Top             =   2370
      Width           =   2955
   End
   Begin VB.FileListBox FLB 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   1830
      Pattern         =   "*.frm;*.bas;*.cls;*.ctl"
      TabIndex        =   4
      Top             =   5280
      Width           =   1935
   End
   Begin VB.ListBox lstChoose 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      ItemData        =   "VB_Editor.frx":2D30
      Left            =   4860
      List            =   "VB_Editor.frx":2D32
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   1950
      Width           =   2385
   End
   Begin VB.FileListBox FLB1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1740
      Pattern         =   "*.vbp"
      TabIndex        =   3
      Top             =   3750
      Width           =   2145
   End
   Begin VB.Label Label2 
      BackColor       =   &H00685758&
      Height          =   1005
      Left            =   30
      TabIndex        =   20
      Top             =   0
      Width           =   18915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nKey                As Long     ' Incremented to ensure unique key
Dim SubKey              As String   ' Key for modules
Dim sKey                As String   ' key for root
Dim selKey              As String   ' key of item selected
Dim cUP                 As Boolean  ' Should next line be indented
Dim cDown               As Boolean  ' should next line be outdented
Dim sDek                As Boolean  ' is it a declaration line
Dim bDirty              As Boolean  ' has the code on this line been edited
Dim ModGot              As Boolean  ' has the Nodule been previously loaded
Dim nInd                As Long     ' Index of last node selected
Private Sub InitLstChoose()
' Fills a ListBox that has checkboxes with
' The mudules in the selected folder
Dim N           As Long

    lstChoose.Clear
    N = FLB1.ListCount
    ' Get the .vbp name to use as
    ' the ROOT of the treeview
    If N > 0 Then
        ' Only has vbp files
        FLB1.ListIndex = 0
        lstChoose.AddItem FLB1.FileName
    Else
        ' Just in case there is no .vbp in directory
        lstChoose.AddItem "VB Project - Unspecified"
    End If
    
    ' Get all the project code files
    For N = 0 To FLB.ListCount - 1
        ' Has all of the .frm, .bas, .cls, .ctl files
        FLB.ListIndex = N
        lstChoose.AddItem FLB.FileName
    Next
End Sub
Private Sub AddRoot(S As String)
Dim nodeA       As Node

    Set nodeA = TRV.Nodes.Add()
    nodeA.Text = S
    sKey = Chr$(nKey) & Chr$(nKey)
    nodeA.Key = sKey
    nodeA.Expanded = True
    nodeA.Bold = True
    nodeA.ForeColor = gnCol(1)
    ' 1st image in the image list
    nodeA.Image = 1
    nKey = nKey + 1
End Sub
Private Sub AddModuleLevel(S As String)
' Creates a top level heading immediately
' below the ROOT for each module
Dim NodeB       As Node
' The text to display on the treeview
Dim sName       As String

    ' Prevent reading it more than once
    If ModGot = False Then
        GetModule (S)
    End If
    SubKey = "B" & Str(nKey)
    sName = S & " - " & sR(2)
    Set NodeB = TRV.Nodes.Add(sKey, tvwChild, SubKey, sName)
    nKey = nKey + 1
    NodeB.Bold = True
    NodeB.ForeColor = gnCol(0)
    NodeB.Image = 8
    NodeB.Expanded = True
    If chkSort.Value = 1 Then
        NodeB.Sorted = True
    End If
    DoProcedures (S)
End Sub
Private Sub DoProcedures(S As String)
' Inserts all of the procedures at the next level
' Indented immediately below the Procedure name
Dim N               As Long
Dim Perr            As Long
Dim Xerr            As Long
Dim Level           As Long
Dim sK(30)          As String
Dim Node(30)        As Node
Dim sPub            As Boolean

    Level = 3
    ' Each item in the Array that contains
    ' Each line to be shown on the treeview
    ' Start at three so that node Index is
    ' the same as the array index
    For N = 3 To UBound(sR)
        S = sR(N)
        
        ' Its a global variable
        If Left$(S, 6) = "Public" Then
            If Left$(S, 10) = "Public Sub" Then
                Level = 3: cUP = True
            
            ElseIf Left$(S, 15) = "Public Property" Then
                cUP = True
            Else
                'its a Public procedure
                sPub = True
            End If
        ElseIf Left$(S, 11) = "Private Sub" Then
            Level = 3: cUP = True
        ElseIf S = "Option Explicit" Then
            Level = 3: cUP = True
        ElseIf S = "End Sub" Then
            Level = 4
        ElseIf Left$(S, 4) = "Else" Then
            ' Line is to be Outdented
            cDown = True
            ' Next Line is to be INdented
            cUP = True
        Else
            ' Check for start/end of nesting
            NestStart (S)
            NestEnd (S)
            ' Check if line is a declaration
            DekLares (S)
        End If
        
        If cDown = True Then
            ' Outdent 1 step
            Level = Level - 1
        End If
        
        ' Level 3 = E , 4 = F, 5 + G et.
        sK(2) = SubKey
        sK(Level) = Chr$(Level + 66) & Str(nKey)
        If Level = 3 Then
            'Node(3).Sorted = True
        End If
        Level = IIf(Level > 2, Level, 3)
        Set Node(Level) = TRV.Nodes.Add(sK(Level - 1), tvwChild, sK(Level), sR(N))
        
        ' See if text contains word 'Error'
        Perr = InStr(S, "Error")
        Xerr = InStr(S, "Exit ")
        
        ' Use conditions to set Colour etc for each node
        If Left$(S, 1) = "'" Then
            Node(Level).ForeColor = gnCol(2)    'Remarks
            Node(Level).Bold = True
            Node(Level).Image = 5
        ElseIf (cUP Or cDown) And Level > 3 Then
             Node(Level).ForeColor = gnCol(6)   'Start or rnd of a nest
        ElseIf sDek = True Then
            Node(Level).ForeColor = gnCol(4)    'Declaration lines
        ElseIf sPub = True Then
            Node(Level).ForeColor = gnCol(3)    'Public variables
        ElseIf Perr > 0 Then
            Node(Level).BackColor = &H80FFFF    'Pale yellow
        ElseIf Xerr > 0 Then
            Node(Level).BackColor = &HC0E0FF    'Pink
        ElseIf Level = 3 Then
            Node(Level).ForeColor = gnCol(1)    'Procedure name
        Else
            Node(Level).ForeColor = gnCol(5)    'The default colour
        End If
            
        ' Make user insertions prominent
        If S = "Inserted" Then
            Node(Level).ForeColor = gnCol(3)
        End If
            
        ' Different Icon for different levels
        If cUP = True Then
            If Level = 3 Then
                Node(Level).Image = 3
            Else
                Node(Level).Image = 2
            End If
            Level = Level + 1
        End If
        
        ' Reset the booleans to false
        '•• Key is Letter = Level & Index ••••••
        nKey = nKey + 1
        cDown = False
        cUP = False
        sDek = False
        sPub = False
    Next
End Sub
Private Sub DekLares(S As String)
' See if line is a declaration
Dim M       As Long
Dim x       As String
Dim L       As Long

    For M = 0 To UBound(sD)
        x = sD(M): L = Len(x)
        If Left$(S, L) = x Then
            sDek = True
            Exit For
        End If
    Next
End Sub
Private Sub NestStart(S As String)
' Is it the Start of a nest
' eg  For, If, Do  etc.
Dim M       As Long
Dim x       As String
Dim L       As Long

    For M = 0 To UBound(sT)
        x = sT(M): L = Len(x)
        'its start of a nest
        If Left$(S, L) = x Then
            'Indent NEXT line by 1 step / tab
            cUP = True
            Exit For
        End If
    Next
End Sub
Private Sub NestEnd(S As String)
' Is it the End of a nest
Dim M       As Long
Dim x       As String
Dim L       As Long

    For M = 0 To UBound(sE)
        x = sE(M): L = Len(x)
        ' its End of a nest
        If Left$(S, L) = x Then
            ' Outdent 1 step
            cDown = True
            Exit For
        End If
    Next
End Sub
Private Sub InitTreeView()
Dim N           As Long
Dim S           As String

    'Looks best with aROOT
    TRV.Visible = True
    TRV.Nodes.Clear
    For N = 0 To lstChoose.ListCount - 1
        lstChoose.ListIndex = N
        S = lstChoose.Text
        If N = 0 Then
            nKey = 1
            AddRoot (S)
        Else
            AddModuleLevel (S)
        End If
    Next
    ' Cannot Edit, Insert, Delete or Save
    ' if loading more than 1 code module
    ' So these buttons are NOT enabled
    If N > 2 Then
        cmdSave.Enabled = False
        cmdInsert(0).Enabled = False
        cmdInsert(1).Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub cmdContinue_Click()
' Process ONly the modules selected
' and the  Projectname
'  - its the root and uses no resources

Dim N           As Long
Dim L           As Long
    
    ' Remove the items not selected
    For N = lstChoose.ListCount - 1 To 1 Step -1
        If lstChoose.Selected(N) = False Then
            lstChoose.RemoveItem N
        End If
    Next
    
    ' Ensure that there is a module to load
    If lstChoose.ListCount < 2 Then
        MsgBox "You MUST check at least ONE module"
        Exit Sub
    End If
    
    ' Uncover the Command Buttons
    ' and the textbox used for editing
    'fraCover.Visible = False
    fraCover.Top = -1660
    Me.Width = 10000
    Me.Height = 10995
    Me.Top = 2000
    cmdExit.Top = 30
    TRV.Left = 90
    InitTreeView
End Sub

Private Sub cmdDelete_Click()
    ' Remove selected node
    TRV.Nodes.Remove (nInd)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdInsert_Click(Index As Integer)
' Inserts a new line below the selected line
' Imeediately below or indented 1 step
Dim sInKey      As String   'Anew key for adding a line
Dim NodeI       As Node
Static k        As Long
Dim N           As Long
Dim U           As Long

    ' Some text MUST be entered Before
    ' an insert can be made.
    If Len(txtEdit.Text) = 0 Or bDirty = False Then
        MsgBox "You MUST enter some text to Insert", _
            vbInformation + vbOKOnly, "INSERT FAILED"
        Exit Sub
    End If
    
    k = k + 1
    sInKey = selKey & Trim(Str(k + 10))
    'sr(index) should match the text of the selected node
    If sR(nInd) = TRV.Nodes.Item(nInd).Text Then
        ' Index and array are synchronised
        ' So it is safe to insert into array
        ' and rebuild treeview
        U = UBound(sR)
        U = U + 1
        ReDim Preserve sR(U)
        For N = U To nInd + 2 Step -1
            sR(N) = sR(N - 1)
        Next
        sR(nInd + 1) = txtEdit.Text
    Else
        MsgBox "Cannot Insert"
        Exit Sub
    End If
    For N = U - 1 To 1 Step -1
         TRV.Nodes.Remove (N)
    Next
    
    ModGot = True
    InitTreeView
    ' Now the inserted line has to be visible
    TRV.Nodes(nInd).EnsureVisible

End Sub

Private Sub cmdOptions_Click()
Dim N               As Long

    frmOptions.Show vbModal
    SaveOptions
    ' If the colours have been changed we
    ' will have to do the nodes again
    If bColors = True Then
        bColors = False
        For N = TRV.Nodes.Count To 1 Step -1
            TRV.Nodes.Remove (N)
        Next
        InitTreeView
        ' Show the last line selected
        TRV.Nodes(nInd).EnsureVisible
    End If
End Sub

Private Sub cmdSave_Click()
' Save the Code including any changes
' that have been made
Dim fNum        As Long
Dim f2Num       As Long

Dim sMName      As String
Dim N           As Long
Dim M           As Long
Dim L           As Long
Dim nD          As Node
Dim S           As String
Dim sL           As String
Dim canStart    As Boolean
Dim iNd         As Long
Dim sUfx        As String

    ' Increment number to uniquely name and
    ' Identify each revision to a module
    nExtNum = nExtNum + 1
    SaveOptions
    
    sUfx = Right$(sR(0), 4)
    L = Len(sR(0))
    sMName = Left$(sR(0), L - 4)
    ' Create a unique BU name for the module
    sMName = sMName & Str(nExtNum) & sUfx
    ' rename mudule with unique identifier
    Name VbPath & sR(0) As VbPath & sMName
        
    N = TRV.Nodes.Count
    fNum = FreeFile
    Open sMName For Input As #fNum
    f2Num = FreeFile
    Open VbPath & sR(0) For Output As #f2Num
        'First copy original file up to Option explicit
        Do Until S = "Option Explicit"
            Line Input #fNum, S
            If S = "Option Explicit" Then
                Exit Do
            End If
            Print #f2Num, S
        Loop
    Close #fNum
    
        ' save code to original filename for module
        For M = 1 To N
            S = TRV.Nodes(M).Text
            ' Only save Code lines
            If S = "Option Explicit" Then
                canStart = True
            End If
            If canStart = True Then
                ' Key starts with E, F, G etc
                ' Indents start with F = 4, G = 8 spaces
                sL = TRV.Nodes(M).Key & "C"
                iNd = (Asc(Left$(sL, 1)) - 69) * 4
                S = Space(iNd) & S
                Print #f2Num, S
            End If
        Next
    Close #f2Num

End Sub

Private Sub DIR_Change()
    FLB.Path = DIR.Path
    ' Listbox used so that user can CHECK
    ' the Modules that are required
    InitLstChoose
    VbPath = FLB.Path & "\"
End Sub

Private Sub DIR_Click()
    FLB.Path = DIR.Path
    InitLstChoose
    VbPath = FLB.Path & "\"
End Sub

Private Sub DRV_Change()
    DIR.Path = DRV.Drive
End Sub

Private Sub Form_Activate()

    ' Centre the form initially
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2 - 600
    ' set treeview to user prefernces
    TRV.Font.Name = sfName
    TRV.Font.Size = nFSize
    TRV.Font.Bold = bBold
    Gradme
    fraCover.ZOrder 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    ' Purely for development to see
    ' actual KeyAscii values
    'Label2.Caption = Str(KeyAscii)
    
    'Enter key pressed and line has been changed
    If KeyAscii = 13 And bDirty = True Then
        ' Transfer the changes to the node
        TRV.Nodes.Item(nInd).Text = txtEdit.Text
        ' Keep the array the same as the nodes
        sR(nInd) = txtEdit.Text
        ' Make the changes prominent
        TRV.Nodes.Item(nInd).ForeColor = gnCol(7)
        ' Changes have been completed
        bDirty = False
        cmdDirty.BackColor = vbWhite
    Else
        ' making changes to txtEdit
        bDirty = True
        cmdDirty.BackColor = vbYellow
    End If
End Sub

Private Sub Form_Load()
   
   ' Initially set drive to Application path
   ' It can then be selected by the user
    DRV.Drive = Left$(App.Path, 3)
    
End Sub

Private Sub Form_Resize()
    If fraCover.Top > 0 Then
        Me.Height = 10300
        Me.Width = 10000
        Exit Sub
    End If
    
    'Make treeview  et. fit the form
    TRV.Width = Me.Width - 300
    TRV.Height = Me.Height - 1440
    txtEdit.Width = Me.Width - 465
    cmdExit.Left = Me.Width - 1335
End Sub

Private Sub TRV_NodeClick(ByVal Node As MSComctlLib.Node)

    ' The Key of the selected node
    selKey = Node.Key
    'lblKey.Caption = selKey
    ' The text of thew selected node
    ' Put text into textbox for editing
    txtEdit.Text = Node.Text
    ' Make index available to ALL procedues
    nInd = Node.Index
    ' Indicates that textbox has not been edited - yet
    bDirty = False
    cmdDirty.BackColor = vbWhite
End Sub
Private Sub Gradme()
Dim R               As Long
Dim G               As Long
Dim B               As Long
Dim R2              As Long
Dim G2              As Long
Dim B2              As Long
Dim Col             As Long
Dim Y               As Long
Dim color1          As Long
Dim color2          As Long

    Me.ScaleMode = vbPixels
    Me.ScaleHeight = 48
    Me.ScaleWidth = 100
    Me.AutoRedraw = True
    Me.DrawStyle = 6
    Me.DrawWidth = 14
    
    color1 = &HC0FFFF       'Olive yellow/green
    color2 = vbBlack
    R = (color1 And 255) And 255
    G = Int(color1 \ 256) And 255
    B = Int(color1 \ 65536) And 255

    For Y = 0 To Me.ScaleHeight
        R2 = Abs(R - 2 * Y)
        G2 = Abs(G - 2 * Y)
        B2 = Abs(B - 2 * Y)
        Me.Line (0, Y)-(Me.ScaleWidth, Y), RGB(R2, G2, B2)
        If Y = 28 Then
            Debug.Print R2, G2, B2
            
        End If
    Next
    Me.AutoRedraw = False
    Me.ScaleMode = vbTwips
End Sub
