' =====================================================================
' Cost-Effectiveness of Gene Therapy in Hemophilia B
' 
' Probabilistic sensitivity analysis, gene therapy vs on-demand treatment.
' Same procedure as PSA_prophylaxis, writing to a separate block of the PSA
' sheet. Produces the right-hand panels of Figure 4.
' 
' NOTE: worksheet names and the values written to the Model_type switch are those
' of the Excel model workbook and are in Chinese. They are reproduced verbatim
' because the macros will not run against the workbook otherwise. Translations:
'   sheet 参数汇总          = Parameters
'   Model_type 基础分析       = base case
'   Model_type 单因素敏感性分析 = one-way sensitivity analysis
'   Model_type 敏感性分析      = probabilistic sensitivity analysis
' =====================================================================

Attribute VB_Name = "PSA_on_demand"
Sub PSA_OD()
    Application.ScreenUpdating = False

    Dim COST_GENE(1000) As Double
    Dim QALYs_GENE(1000) As Double
    Dim COST_OD(1000) As Double
    Dim QALYs_OD(1000) As Double
    Dim current_deterministic_QALYs_OD(1000) As Double
    Dim current_deterministic_costs_OD(1000) As Double

    Sheets("参数汇总").Activate
    Range("Model_type").Value = "敏感性分析"

    For i = 1 To 1000
        Application.StatusBar = "Simulation " & i & " of " & 1000
        Calculate
        COST_GENE(i) = Range("COST_GENE").Value
        QALYs_GENE(i) = Range("QALYs_GENE").Value
        COST_OD(i) = Range("COST_OD").Value
        QALYs_OD(i) = Range("QALYs_OD").Value
        current_deterministic_costs_OD(i) = COST_GENE(i) - COST_OD(i)
        current_deterministic_QALYs_OD(i) = QALYs_GENE(i) - QALYs_OD(i)
    Next i

    Sheets("PSA").Select
    For i = 1 To 1000
        Cells(12 + i, 42).Value = COST_GENE(i)                        ' AP列
        Cells(12 + i, 43).Value = QALYs_GENE(i)                       ' AQ列
        Cells(12 + i, 44).Value = COST_OD(i)                          ' AR列
        Cells(12 + i, 45).Value = QALYs_OD(i)                         ' AS列
        Cells(12 + i, 47).Value = current_deterministic_costs_OD(i)   ' AU列
        Cells(12 + i, 48).Value = current_deterministic_QALYs_OD(i)   ' AV列
    Next i

 Sheets("PSA").Select
For i = 1 To 57
    Range("wtp").Value = Cells(12 + i, 51).Value
    Calculate
    Cells(12 + i, 52).Value = Range("prob_ce_OD").Value   ' 改成AZ列(52)，跟M/AZ对应上
Next i

    Sheets("参数汇总").Activate
    Range("Model_type").Value = "基础分析"

    Application.StatusBar = "PSA_OD complete"
    Range("A1").Select
End Sub
