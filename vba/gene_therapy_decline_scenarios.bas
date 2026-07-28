' =====================================================================
' Cost-Effectiveness of Gene Therapy in Hemophilia B
' 
' Factor IX activity decline scenarios.
' Sweeps the cycle at which activity crosses each band boundary and
' records the resulting costs and QALYs. Produces Supplementary Table S18.
' 
' NOTE: worksheet names and the values written to the Model_type switch
' are those of the Excel model workbook and are in Chinese. They are
' reproduced verbatim because the macros will not run against the
' workbook otherwise. Translations:
'   sheet 参数汇总        = Parameters
'   sheet OWSA / PSA / Result are already in English
'   Model_type 基础分析       = base case
'   Model_type 单因素敏感性分析 = one-way sensitivity analysis
'   Model_type 敏感性分析      = probabilistic sensitivity analysis
' =====================================================================

Attribute VB_Name = "gene_therapy_decline_scenarios"
Sub SensitivityAnalysis()
Attribute SensitivityAnalysis.VB_Description = "for the degrade of gene thearpy"
Attribute SensitivityAnalysis.VB_ProcData.VB_Invoke_Func = " \n14"
    ' Declare variables
    Dim ws As Worksheet
    Dim i As Long
    Dim startVal As Long
    Dim endVal As Long
    Dim stepVal As Long
    Dim outputRow As Long
    
    On Error GoTo ErrorHandler ' Enable error handling
    
    ' Set the worksheet (with error checking)
    Set ws = Nothing
    For Each ws In ThisWorkbook.Sheets
        If ws.Name Like "*Trace_Gen_sensitivity analysis*" Then
            Set ws = ThisWorkbook.Sheets(ws.Name)
            Exit For
        End If
    Next ws
    
    ' Check if worksheet was found
    If ws Is Nothing Then
        MsgBox "Error: Worksheet 'Trace_Gen_sensitivity analysis' not found!", vbCritical
        Exit Sub
    End If
    
    ' Initialize parameters
    startVal = 1
    endVal = 4300
    stepVal = 1
    outputRow = 5 ' Starting row for output
    
    ' Clear previous results (with error handling)
    On Error Resume Next
    ws.Range("K5:P" & ws.Rows.Count).ClearContents
    On Error GoTo ErrorHandler
    
    ' Disable screen updates for speed
    Application.ScreenUpdating = False
    
    ' Loop through values
    For i = startVal To endVal Step stepVal
        ' Set value in I3
        ws.Range("I3").Value = i
        
        ' Force calculation (handle manual/auto modes)
        If Application.Calculation = xlCalculationAutomatic Then
            Application.Calculate
        Else
            ws.Calculate
        End If
        
        ' Record results (with explicit checks)
        With ws
            If Not IsError(.Range("B16").Value) Then .Cells(outputRow, "K").Value = .Range("B16").Value
            If Not IsError(.Range("C16").Value) Then .Cells(outputRow, "L").Value = .Range("C16").Value
            If Not IsError(.Range("B20").Value) Then .Cells(outputRow, "O").Value = .Range("B20").Value
            If Not IsError(.Range("C20").Value) Then .Cells(outputRow, "P").Value = .Range("C20").Value
            .Cells(outputRow, "J").Value = i ' Store input value
        End With
        
        outputRow = outputRow + 1
        
        ' Status update every 10 iterations (adjustable)
        If i Mod 10 = 0 Then
            Application.StatusBar = "Processing: " & i & " of " & endVal
            DoEvents
        End If
    Next i
    
    ' Clean up
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    MsgBox "Analysis completed for I3 = " & startVal & " to " & endVal, vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Occurred at I3 = " & i, vbCritical
End Sub
