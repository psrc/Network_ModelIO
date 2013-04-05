Imports ESRI.ArcGIS.GeoDatabaseUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework

Public Class clsCalculator
    Implements ICalculatorCallback
    Private pStatBar As IStatusBar
    Private pProbar As IStepProgressor
    '++ Application ref for statusbar
    Private pAppRef As AppRef
    Private fso As Object
    Private f1 As Object
    Private file_nme As String
    Public Sub New()
        '++ Create a file for logging errors
        file_nme = "c:\temp\CalcError.log"
        fso = CreateObject("Scripting.FileSystemObject")
        f1 = fso.CreateTextFile(file_nme, True)
        '++ Create Progress bar
        pAppRef = New AppRef
        pStatBar = pAppRef.StatusBar
        pProbar = pStatBar.ProgressBar
        pProbar.MinRange = 0
        pProbar.Position = 0
        '++ MaxRange is reset later, but is set
        '++ here to allow stepvalue to be set
        pProbar.MaxRange = 10
        pProbar.StepValue = 1
        pProbar.Show()
    End Sub


    Public Function CalculatorError(ByVal rowID As Integer, ByVal bHasOID As Boolean, ByVal errorType As ESRI.ArcGIS.GeoDatabaseUI.esriCalculatorErrorType, ByVal bShowPrompt As Boolean, ByVal errorMsg As String) As Boolean Implements ESRI.ArcGIS.GeoDatabaseUI.ICalculatorCallback.CalculatorError
        '++ Get a desc for the error
        Dim Desc As String
        Desc = Get_Desc(errorType)
        '++ write error details to log file
        f1.WriteLine("Row ID :" & rowID & "; " & "Error TYPE :" & errorType & "; " & Desc & "; " & errorMsg)
        '++ If certain error types are encountered,
        '++ setting this to true will halt the calculation
        If errorType = 0 Or errorType = 1 Or errorType = 2 Or errorType = 3 Then
            CalculatorError = True
            f1.WriteLine("Failed to complete calculation.")
        End If
    End Function

    Public Sub CalculatorWarning(ByVal rowID As Integer, ByVal bHasOID As Boolean, ByVal errorType As ESRI.ArcGIS.GeoDatabaseUI.esriCalculatorErrorType, ByVal errorMsg As String) Implements ESRI.ArcGIS.GeoDatabaseUI.ICalculatorCallback.CalculatorWarning
        '++ Custom warning
        pProbar.Hide()
        pStatBar.Message(0) = "Warning:- error at " & rowID
        pStatBar.Message(0) = ""
        pProbar.Show()
    End Sub

    Public Function Status(ByVal rowsWritten As Integer, ByVal lastStatus As Boolean) As Boolean Implements ESRI.ArcGIS.GeoDatabaseUI.ICalculatorCallback.Status
        If rowsWritten < 1 Then
            '++ cancel the calculation
            Status = True
        Else
            '++ Update the progress bar
            pProbar.Step()
            If lastStatus = True Then
                pProbar.Hide()
                pProbar.Position = 0
            End If
        End If
        '++ Add the rwoWritten to the log file
        f1.WriteLine("Processed row: " & rowsWritten)
    End Function
    Private Sub Finalize()
        '++ De_ref the Status/Progress bar objects
        '++ and close the log file
        pProbar.Hide()
        pProbar = Nothing
        pStatBar = Nothing
        f1.Close()
    End Sub
    Private Function Get_Desc(ByVal ET As esriCalculatorErrorType) As String
        Select Case ET
            Case 0
                Get_Desc = "No script environment(VB/A) installed"
            Case 1
                Get_Desc = "Error parsing expression"
            Case 2
                Get_Desc = "Error running VB/A code"
            Case 3
                Get_Desc = "Empty value error (e.g. calculated value is invalid for the specified field)."
            Case 4
                Get_Desc = "Error putting calculated value into the specified field"
            Case 5
                Get_Desc = "Error storing calculated value into the specified field"
            Case 6
                Get_Desc = "All calculated values are invalid"
        End Select
    End Function
    Public Sub Get_Rowset(ByVal rowct As Long)
        pProbar.MaxRange = rowct

    End Sub
End Class
