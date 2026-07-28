' =====================================================================
' Cost-Effectiveness of Gene Therapy in Hemophilia B
' 
' Probabilistic sensitivity analysis, gene therapy vs prophylaxis.
' Draws 1,000 parameter sets, records paired costs and QALYs for both strategies,
' then evaluates the probability of cost-effectiveness at 57 willingness-to-pay
' thresholds. Produces the left-hand panels of Figure 4.
' 
' NOTE: worksheet names and the values written to the Model_type switch are those
' of the Excel model workbook and are in Chinese. They are reproduced verbatim
' because the macros will not run against the workbook otherwise. Translations:
'   sheet 参数汇总          = Parameters
'   Model_type 基础分析       = base case
'   Model_type 单因素敏感性分析 = one-way sensitivity analysis
'   Model_type 敏感性分析      = probabilistic sensitivity analysis
' =====================================================================

Attribute VB_Name = "PSA_prophylaxis"
Sub PSA_PH()
    Application.ScreenUpdating = False
    
    Dim COST_GENE(1000) As Double
    Dim QALYs_GENE(1000) As Double
    Dim COST_PH(1000) As Double
    Dim QALYs_PH(1000) As Double
    Dim current_deterministic_QALYs_PH(1000) As Double
    Dim current_deterministic_costs_PH(1000) As Double
    
    Sheets("参数汇总").Activate
    Range("Model_type").Value = "敏感性分析"
    
    For i = 1 To 1000
        Application.StatusBar = "Simulation " & i & " of " & 1000
        Calculate
        COST_GENE(i) = Range("COST_GENE").Value
        QALYs_GENE(i) = Range("QALYs_GENE").Value
        COST_PH(i) = Range("COST_PH").Value
        QALYs_PH(i) = Range("QALYs_PH").Value
        current_deterministic_costs_PH(i) = COST_GENE(i) - COST_PH(i)
        current_deterministic_QALYs_PH(i) = QALYs_GENE(i) - QALYs_PH(i)
    Next i
    
    Sheets("PSA").Select
    For i = 1 To 1000
        Cells(12 + i, 3).Value = COST_GENE(i)
        Cells(12 + i, 4).Value = QALYs_GENE(i)
        Cells(12 + i, 5).Value = COST_PH(i)
        Cells(12 + i, 6).Value = QALYs_PH(i)
        Cells(12 + i, 8).Value = current_deterministic_costs_PH(i)
        Cells(12 + i, 9).Value = current_deterministic_QALYs_PH(i)
    Next i
    
   Sheets("PSA").Select
For i = 1 To 57
    Range("wtp").Value = Cells(12 + i, 12).Value
    Calculate
    Cells(12 + i, 13).Value = Range("prob_ce").Value
Next i
    
     Sheets("参数汇总").Activate
    Range("Model_type").Value = "基础分析"
    
    Application.StatusBar = "PSA complete"
    Range("A1").Select
End Sub
