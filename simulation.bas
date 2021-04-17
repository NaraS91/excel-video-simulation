Attribute VB_Name = "Module1"

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Sub PasteFrames()
Attribute PasteFrames.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PasteFrames Macro
'

'
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim pathToFrames As String, minFrameTime As Integer
    
'   change pathToFrames to store the path to folder where the frames are stored
'   minFrameTime decides on the minimum time a frame is shown on screen
'   assign 0 for the fastest render
'   or delete the section with events and ticks in the last for for maximum efficiency
    pathToFrames = "D:\Documents\badApple\5"
    minFrameTime = 0
    
    Set oFolder = oFSO.GetFolder(pathToFrames)

    MsgBox "Press ok to start loading frames into memory. It might take a while."

    Dim Values() As Variant
    ReDim Values(oFolder.Files.Count)

    For i = 1 To oFolder.Files.Count
        Values(i) = LoadFrames(pathToFrames & "\frame" & i & ".txt")
    Next i
    
    Dim rowsCount As Integer, columnsCount As Integer
    rowsCount = UBound(Values(1), 1) - LBound(Values(1), 1) + 1
    columnsCount = UBound(Values(1), 2) - LBound(Values(1), 2) + 1
    
    MsgBox "frames finished loading to memory"
    MsgBox "simulation will start after you press ok. To stop it earlier press ctrl + break"
    
    Dim now As Long, finish As Long
    
    For i = 1 To oFolder.Files.Count
        finish = GetTickCount() + frameTime
        Range(Cells(1, 1), Cells(rowsCount, columnsCount)) = Values(i)
        Do
            DoEvents
            now = GetTickCount()
        Loop Until now >= finish
    Next i
    
End Sub

Function LoadFrames(filePath As String) As Variant
    Open filePath For Input As #1
    Content = Input(LOF(1), 1)
    Close #1
    
    Dim lines() As String, rowsCount As Integer, columnsCount As Integer
    lines = Split(Content, vbLf)
    rowsCount = UBound(lines)
    columnsCount = UBound(Split(lines(0)))
    Dim Values() As Variant
    ReDim Values(rowsCount, columnsCount)
    
    For i = 0 To UBound(lines)
        Dim numbersStrings() As String
        numbersStrings = Split(lines(i))
        
        For j = 0 To UBound(numbersStrings)
            Values(i, j) = CByte(CInt(numbersStrings(j)))
        Next j
    Next i
    LoadFrames = Values
End Function
