Attribute VB_Name = "modVBEd"
Option Explicit
Public sR()         As String   'Array to hold contents of modeul
Public sfName       As String   'Selected font name
Public nFSize       As Long     'Selected fontsize
Public bBold        As Boolean  'Bold / Not bold selected
Public bColors      As Boolean  'Colours have been changed
Public fName        As String   'Filename to save options to
Public sFile        As String   'Name of loaded module
Public nExtNum      As Long     'Number to append to bu filename
Public gnCol(10)    As Long     'Text colours selected
Public sT()                     'Variants to assign to an array
Public sE()
Public sD()
Public VbPath       As String   ' The path selected to work with

Public Sub SaveOptions()
' Save the changes to the options
' Options that need to be saved are
' Fontname, Fontsize, Bold y/n
' numbers of the 7 font colours
Dim fNum        As Long
Dim N           As Long
    
    fNum = FreeFile
    Open fName For Output As #fNum
        Print #fNum, sfName
        Print #fNum, Str(nFSize)
        Print #fNum, bBold
        For N = 0 To 7
            Print #fNum, gnCol(N)
        Next
        Print #fNum, nExtNum
    Close #fNum
End Sub
Private Sub GetOptions()
' Recall the most recent option settings
Dim fNum        As Long
Dim N           As Long
    
    On Error GoTo FileReadError
    fNum = FreeFile
    fName = App.Path & "\Options.txt"
    Open fName For Input As #fNum
        Input #fNum, sfName
        Input #fNum, nFSize
        Input #fNum, bBold
        For N = 0 To 7
            Input #fNum, gnCol(N)
        Next
        Input #fNum, nExtNum
    Close #fNum
    
Exit Sub
FileReadError:
    Close #fNum
    MsgBox "No Options file found"
End Sub
Private Sub SetArrays()
    
    ' Lazy way to put several items into an array
    ' Start of Nests
    ' These have a space after them to avoid
    ' a hit on words like Forecolor, Double etc.
    sT = Array("For ", "If ", "Do ", "Open ")
    
    ' End of nests
    sE = Array("Next ", "End If", "Loop ", "Close ", "End Property")
    
    ' Procedure level declarations
    sD = Array("Dim", "Static", "ReDim")
End Sub
Private Sub Main()
    
    SetArrays
    GetOptions
    frmMain.Show
End Sub

'• This module returns just the code
'• for the .bas, .frm and .cls and .ctl files
'• That are selected from the current Application Path
Public Sub GetModule(sF As String)
Dim fN          As Long
Dim N           As Long
Dim S           As String
Dim sTart       As Boolean
ReDim sR(3000)
    
    On Error GoTo FileReadError
    
    fN = FreeFile
    sR(0) = sF
    'sF is the file name of each module selected
    sFile = VbPath & sF
    N = 2
    ' Read ALL of the code lines
    ' Note sR(1) is project name
    ' sR(2) will be first module name
    
    Open sFile For Input As #fN
        Do While Not EOF(fN)
            Line Input #fN, S
            S = Trim(S)
            If Left$(S, 13) = "Begin VB.Form" Then
                sR(N) = Mid$(S, 14, 12)
            End If
            '• Skip over lines that are about controls
            If S = "Option Explicit" Then
                sTart = True
            End If
            
            '• Skip Blank lines
            If sTart And Len(S) > 1 Then
                N = N + 1
                sR(N) = S 'Option Explicit should be sR(3)
            End If
            If Left$(sF, 3) = "cls" Then
                sTart = True
            End If
        Loop
    Close #fN
    ReDim Preserve sR(N)
    Exit Sub
    
FileReadError:
    MsgBox "Error retrieving File", vbCritical + vbOKOnly, "FILE ERROR"
End Sub
