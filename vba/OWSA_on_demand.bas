' =====================================================================
' Cost-Effectiveness of Gene Therapy in Hemophilia B
' 
' One-way sensitivity analysis, gene therapy vs recombinant factor IX on-demand
' treatment. Same procedure as OWSA_prophylaxis. Produces Supplementary Table S17.
' 
' NOTE: worksheet names and the values written to the Model_type switch are those
' of the Excel model workbook and are in Chinese. They are reproduced verbatim
' because the macros will not run against the workbook otherwise. Translations:
'   sheet 参数汇总          = Parameters
'   Model_type 基础分析       = base case
'   Model_type 单因素敏感性分析 = one-way sensitivity analysis
'   Model_type 敏感性分析      = probabilistic sensitivity analysis
' =====================================================================

Attribute VB_Name = "OWSA_on_demand"
Sub OWSA_OD()

Application.ScreenUpdating = False

Sheets("参数汇总").Activate

Sheets("参数汇总").Range("Model_type").Value = "单因素敏感性分析"

Application.StatusBar = "Running OWSA..."


'Begin the analysis

Dim loca1, loca2, loca3  As String

For i = 0 To 51

loca1 = Sheets("OWSA").Range("OWSA_name").Offset(i, 0).Value
loca2 = Sheets("OWSA").Range("OWSA_nameL").Offset(i, 0).Value
loca3 = Sheets("OWSA").Range("OWSA_nameH").Offset(i, 0).Value

Set range1 = Range(loca1)
Set range2 = Range(loca2)
Set range3 = Range(loca3)

range1.Value = range2.Value

Calculate

Sheets("OWSA").Range("owsa_baseline_results_costs_L_OD").Offset((1 + i), 0).Value = Sheets("Result").Range("current_deterministic_costs_OD").Value
Sheets("OWSA").Range("owsa_baseline_results_QALYs_L_OD").Offset((1 + i), 0).Value = Sheets("Result").Range("current_deterministic_QALYs_OD").Value
Sheets("OWSA").Range("owsa_baseline_results_ICER_L_OD").Offset((1 + i), 0).Value = Sheets("Result").Range("current_deterministic_ICER_OD").Value
Sheets("OWSA").Range("owsa_baseline_results_NMB_L_OD").Offset((1 + i), 0).Value = Sheets("Result").Range("current_deterministic_NMB_OD").Value

'Clears the lower variation restoring the deterministic inputs for the upper variation

Sheets("参数汇总").Activate
Range("F13:F108").ClearContents

Calculate

range1.Value = range3.Value

Calculate

Sheets("OWSA").Range("owsa_baseline_results_costs_H_OD").Offset((1 + i), 0).Value = Sheets("Result").Range("current_deterministic_costs_OD").Value
Sheets("OWSA").Range("owsa_baseline_results_QALYs_H_OD").Offset((1 + i), 0).Value = Sheets("Result").Range("current_deterministic_QALYs_OD").Value
Sheets("OWSA").Range("owsa_baseline_results_ICER_H_OD").Offset((1 + i), 0).Value = Sheets("Result").Range("current_deterministic_ICER_OD").Value
Sheets("OWSA").Range("owsa_baseline_results_NMB_H_OD").Offset((1 + i), 0).Value = Sheets("Result").Range("current_deterministic_NMB_OD").Value

Calculate

'Clear DSA inputs in the model parameters sheet

Application.ScreenUpdating = False


Sheets("参数汇总").Activate
Range("F13:F108").ClearContents
        
Application.ScreenUpdating = False


Next i

'Set DSA switch off

Range("model_type") = "基础分析"

Sheets("参数汇总").Activate

Application.StatusBar = "DSA complete"

Range("A1").Select



End Sub
