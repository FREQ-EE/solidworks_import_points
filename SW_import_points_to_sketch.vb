Dim swApp As SldWorks.SldWorks
Dim swModel As ModelDoc2
Dim swSketchMgr As SketchManager
Dim filePath As String
Dim line As String
Dim fileNumber As Integer

Sub main()

    ' Set the active SolidWorks application and model
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    ' Ensure there is an open part document
    If swModel Is Nothing Or swModel.GetType <> swDocPART Then
        MsgBox "Please open a part document to run this macro."
        Exit Sub
    End If

    ' Set the path to the CSV file
    filePath = "FILE\PATH\HERE"  ' Update this path to the actual file location

    ' Read and parse the CSV file
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber

    ' Begin a new 3D sketch
    Set swSketchMgr = swModel.SketchManager
    swSketchMgr.Insert3DSketch True

    ' Read each line and extract the coordinates to create a point
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        ' Skip header row
        If InStr(line, "Bone Name") = 0 Then
            Dim splitLine() As String
            splitLine = Split(line, ",")
            Dim x As Double, y As Double, z As Double
            x = CDbl(splitLine(1))
            y = CDbl(splitLine(2))
            z = CDbl(splitLine(3))
            
            ' Create a point at the specified coordinates
            swSketchMgr.CreatePoint x, y, z
        End If
    Loop

    ' Close the CSV file
    Close #fileNumber

    ' End the 3D sketch
    swSketchMgr.Insert3DSketch True

    MsgBox "Points created from CSV log."

End Sub