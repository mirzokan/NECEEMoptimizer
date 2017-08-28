Option Explicit

Sub ImportWorksheets(file_path As Range, file_list As Variant)
 ' This macro will import multiple files into this workbook
    Dim i As Long

    For i = 0 To UBound(file_list)
        Call ImportSheet(file_path, file_list(i), i + 1)
        Application.Wait (Now + #12:00:01 AM#)
    Next i    
End Sub

Sub ImportSheet(file_path As Range, filename As Variant, sheet_num As Long)
 ' This macro will import a file into this workbook
    Dim source As Variant

    ' Application.ScreenUpdating = False
    ' Application.Calculation = xlCalculationManual
    ' Application.DisplayAlerts = False
      
    sh_antemp.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Worksheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Worksheets.Count).Name = "DataSet_" & sheet_num
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H2").Value = filename
    
    On Error GoTo fail_files
    Set source = Application.Workbooks.Open(file_path.Value & filename, ReadOnly:=True)
    
    On Error GoTo 0
    Application.DisplayAlerts = False
    'Copy from source sheet
    Windows(filename).Activate
    source.Sheets(1).Range("A1:E" & source.Sheets(1).Range("A14").End(xlDown).Row).Select
    Selection.Copy
          
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Activate
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("A1").PasteSpecial xlPasteValues

    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("A1").Calculate
             
    Windows(filename).Activate
    ActiveWorkbook.Close SaveChanges:=False
    sh_data.Activate
    sh_data.Calculate
    
    Application.DisplayAlerts = True
    
    clear_template (sheet_num)
    copy_global (sheet_num)
    adjust_template (sheet_num)
    
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("A1").Calculate


    ' Application.ScreenUpdating = True
    ' Application.Calculation = xlCalculationAutomatic
    ' Application.DisplayAlerts = True     
    Exit Sub

    
    ' In case of error
    fail_files:
    MsgBox "Problem opening files. Please verify the file path."
    
    ' Application.ScreenUpdating = True
    ' Application.Calculation = xlCalculationAutomatic
    ' Application.DisplayAlerts = True 
End Sub

Sub clear_template(sheet_num As Long)
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("I52:I" & sh_data.Range("I52").End(xlDown).Row).Clear
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("J50:J" & sh_data.Range("J50").End(xlDown).Row).Clear
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("L52:L" & sh_data.Range("L52").End(xlDown).Row).Clear
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("M51:M" & sh_data.Range("M51").End(xlDown).Row).Clear
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("Q51:Q" & sh_data.Range("Q51").End(xlDown).Row).Clear
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("R50:R" & sh_data.Range("R50").End(xlDown).Row).Clear
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("S50:S" & sh_data.Range("S50").End(xlDown).Row).Clear

    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("V51:V" & sh_data.Range("V51").End(xlDown).Row).Clear
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("W51:W" & sh_data.Range("W51").End(xlDown).Row).Clear

    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("Z50:Z" & sh_data.Range("Z50").End(xlDown).Row).Clear

    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("AA51:AI" & sh_data.Range("AA51").End(xlDown).Row).Clear
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("AK50:AS" & sh_data.Range("A50").End(xlDown).Row).Clear
End Sub

Sub copy_global(sheet_num As Long)
    'Copy Global Values
    'L total
    sh_data.Range("exp_conc0_l").Copy Destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H34")
    'T total
    sh_data.Range("exp_conc0_t").Copy Destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H35")
    'Quantum yield ratio
    sh_data.Range("quantyield_ratio").Copy Destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H36")
    'Work Area start
    sh_data.Range("ns_wr_start").Copy Destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H37")
    'Background region 1
    sh_data.Range("ns_bg_start").Copy Destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H38")
    'Background region 2
    sh_data.Range("ns_bg_end").Copy Destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H39")
    'Work Area end
    sh_data.Range("ns_wr_end").Copy Destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H40")
    'Sigma
    sh_data.Range("override_s2n").Copy Destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H41")
    'Propagation Time
    sh_data.Range("ns_exp_prop_time").Copy Destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("G50")
    'Peak width L
    sh_data.Range("override_peakwidth_l").Copy Destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H42")
    'Peak width C
    sh_data.Range("override_peakwidth_c").Copy Destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H43")
End Sub

Sub adjust_template(sheet_num As Long)
    Dim total_points As Long, sampling_rate As Long, segment_start As Long, segment_end As Long, segment_point_count As Long, total_bg_interval As Long, max_a_points As Long, max_c_points As Long
    Dim cur_sheet As Worksheet

    Set cur_sheet = ThisWorkbook.Sheets("DataSet_" & sheet_num)

    cur_sheet.Range("H34:J43").Calculate

    total_points = cur_sheet.Range("B9").Value
    sampling_rate = cur_sheet.Range("B8").Value
    segment_start = cur_sheet.Range("J37").Value
    segment_end = cur_sheet.Range("J40").Value

    segment_point_count = (segment_end - segment_start) * sampling_rate
    total_bg_interval = (cur_sheet.Range("J38").Value + cur_sheet.Range("J39").Value) * sampling_rate

    'Fill out dynamic ranges
    'time, s column
    cur_sheet.Range("I51").Copy Destination:=cur_sheet.Range("I52:I" & total_points + 49)
    'orig column
    cur_sheet.Range("J50:J" & total_points + 49).FormulaArray = "=OFFSET($A$1,13,,$B$9)"
    'index column
    cur_sheet.Range("L51").Copy Destination:=cur_sheet.Range("L52:L" & segment_point_count + 49)
    'time min column
    cur_sheet.Range("M50").Copy Destination:=cur_sheet.Range("M51:M" & total_points + 49)
    'time_window column
    cur_sheet.Range("R50:R" & segment_point_count + 49).FormulaArray = "=OFFSET($I$50,($J$37+$H$50)*$Q$41,,($J$40-$J$37)*$Q$41+1)"
    'time_window_min column
    cur_sheet.Range("Q50").Copy Destination:=cur_sheet.Range("Q51:Q" & segment_point_count + 49)
    'RFU_window column
    cur_sheet.Range("S50:S" & segment_point_count + 49).FormulaArray = "=OFFSET($J$50,($J$37+$H$50)*$Q$41,,($J$40-$J$37)*$Q$41+1)"
    'bg_time
    cur_sheet.Range("V50").Copy Destination:=cur_sheet.Range("V51:V" & total_bg_interval + 49)
    'bg_RFU
    cur_sheet.Range("W50").Copy Destination:=cur_sheet.Range("W51:W" & total_bg_interval + 49)
    'RFU_corr
    cur_sheet.Range("Z50:Z" & segment_point_count + 49).FormulaArray = "=RFU_window-(time_window*$U$35+$V$35)"
    'RFU_rev
    cur_sheet.Range("AA50").Copy Destination:=cur_sheet.Range("AA51:AA" & segment_point_count + 49)
    'Divergency
    cur_sheet.Range("AC50").Copy Destination:=cur_sheet.Range("AC51:AC" & segment_point_count + 49)
    'max1
    cur_sheet.Range("AD50").Copy Destination:=cur_sheet.Range("AD51:AD" & segment_point_count + 49)
    'tc,st
    cur_sheet.Range("AE50").Copy Destination:=cur_sheet.Range("AE51:AE" & segment_point_count + 49)
    'Div_rev
    cur_sheet.Range("AG50").Copy Destination:=cur_sheet.Range("AG51:AG" & segment_point_count + 49)
    'max2
    cur_sheet.Range("AH50").Copy Destination:=cur_sheet.Range("AH51:AH" & segment_point_count + 49)
    'ta,end
    cur_sheet.Range("AI50").Copy Destination:=cur_sheet.Range("AI51:AI" & segment_point_count + 49)


    'max1_RFU
    cur_sheet.Range("AK50:AK" & (segment_point_count / 2) + 49).FormulaArray = "=OFFSET($Z$50,($AD$34-$J$37)*$Q$41-$H$42*$AD$38+1,,$H$42*$AD$38*2)"
    'max1_time
    cur_sheet.Range("AL50:AL" & (segment_point_count / 2) + 49).FormulaArray = "=OFFSET($R$50,($AD$34-$J$37)*$Q$41-$H$42*$AD$38+1,,$H$42*$AD$38*2)"
    'max1_time2
    cur_sheet.Range("AM50:AM" & (segment_point_count / 2) + 49).FormulaArray = "=max1_time^2"
    'max1_fit
    cur_sheet.Range("AN50:AN" & (segment_point_count / 2) + 49).FormulaArray = "=AL50:AL350^2*$AK$35+AL50:AL350*$AL$35+$AM$35"

    'max2_RFU
    cur_sheet.Range("AP50:AP" & (segment_point_count / 2) + 49).FormulaArray = "=OFFSET($Z$50,($AH$34-$J$37)*$Q$41-$H$43*$AH$38+1,,$H$43*$AH$38*2)"
    'max2_time
    cur_sheet.Range("AQ50:AQ" & (segment_point_count / 2) + 49).FormulaArray = "=OFFSET($R$50,($AH$34-$J$37)*$Q$41-$H$43*$AH$38+1,,$H$43*$AH$38*2)"
    'max2_time2
    cur_sheet.Range("AR50:AR" & (segment_point_count / 2) + 49).FormulaArray = "=max2_time^2"
    'max2_fit
    cur_sheet.Range("AS50:AS" & (segment_point_count / 2) + 49).FormulaArray = "=AQ50:AQ350^2*$AP$35+AQ50:AQ350*$AQ$35+$AR$35"


    'Update charts
    cur_sheet.ChartObjects("Chart 1").Activate
    ActiveChart.SeriesCollection(1).XValues = cur_sheet.Range("M50:M" & total_points + 49)
    ActiveChart.SeriesCollection(1).Values = cur_sheet.Range("J50:J" & total_points + 49)
    cur_sheet.ChartObjects("Chart 2").Activate
    ActiveChart.SeriesCollection(1).XValues = cur_sheet.Range("R50:R" & segment_point_count + 49)
    ActiveChart.SeriesCollection(1).Values = cur_sheet.Range("Z50:Z" & segment_point_count + 49)
    ActiveChart.Axes(xlCategory).MinimumScale = segment_start - (segment_end - segment_start) * 0.1
    ActiveChart.Axes(xlCategory).MaximumScale = segment_end + (segment_end - segment_start) * 0.1


    cur_sheet.ChartObjects("Chart 3").Activate
    ActiveChart.SeriesCollection(1).XValues = cur_sheet.Range("Q50:Q" & segment_point_count + 49)
    ActiveChart.SeriesCollection(1).Values = cur_sheet.Range("Z50:Z" & segment_point_count + 49)
    ActiveChart.Axes(xlCategory).MinimumScale = (segment_start / 60)
    ActiveChart.Axes(xlCategory).MaximumScale = (segment_end / 60)


    cur_sheet.ChartObjects("Chart 4").Activate
    ActiveChart.SeriesCollection(1).XValues = cur_sheet.Range("AL50:AL" & (segment_point_count / 2) + 49)
    ActiveChart.SeriesCollection(1).Values = cur_sheet.Range("AK50:AK" & (segment_point_count / 2) + 49)
    ActiveChart.SeriesCollection(2).XValues = cur_sheet.Range("AL50:AL" & (segment_point_count / 2) + 49)
    ActiveChart.SeriesCollection(2).Values = cur_sheet.Range("AN50:AN" & (segment_point_count / 2) + 49)


    cur_sheet.ChartObjects("Chart 5").Activate
    ActiveChart.SeriesCollection(1).XValues = cur_sheet.Range("AQ50:AQ" & (segment_point_count / 2) + 49)
    ActiveChart.SeriesCollection(1).Values = cur_sheet.Range("AP50:AP" & (segment_point_count / 2) + 49)
    ActiveChart.SeriesCollection(2).XValues = cur_sheet.Range("AQ50:AQ" & (segment_point_count / 2) + 49)
    ActiveChart.SeriesCollection(2).Values = cur_sheet.Range("AS50:AS" & (segment_point_count / 2) + 49)

    cur_sheet.ChartObjects("Chart 6").Activate
    ActiveChart.SeriesCollection(1).XValues = cur_sheet.Range("R50:R" & segment_point_count + 49)
    ActiveChart.SeriesCollection(1).Values = cur_sheet.Range("AC50:AC" & segment_point_count + 49)
    ActiveChart.Axes(xlCategory).MinimumScale = (segment_start)
    ActiveChart.Axes(xlCategory).MaximumScale = (segment_end)
End Sub

Sub print_filelist(file_list As Variant)
 'This Sub prints the list of files in a folder in the result table
    Dim i As Integer, i_av As Integer, ii As Integer, copy_cols As Integer
    copy_cols = 8

    For i = 0 To UBound(file_list)
        'Filenumber
        sh_naaped.Cells(i + 2, "A").Value = i + 1
        'Filename
        sh_naaped.Cells(i + 2, "B").Value = file_list(i)
        'Kd
        sh_naaped.Cells(i + 2, "C").Value = "=DataSet_" & i + 1 & "!H44"
        'koff
        sh_naaped.Cells(i + 2, "D").Value = "=DataSet_" & i + 1 & "!H45"
        'Elution time complex
        sh_naaped.Cells(i + 2, "E").Value = "=DataSet_" & i + 1 & "!AD40"
        'Elution time Ligand
        sh_naaped.Cells(i + 2, "F").Value = "=DataSet_" & i + 1 & "!AH40"
        'Noise
        sh_naaped.Cells(i + 2, "G").Value = "=DataSet_" & i + 1 & "!V37"
        'Peak Ligand
        sh_naaped.Cells(i + 2, "H").Value = "=DataSet_" & i + 1 & "!Z35"
        'Peak Complex
        sh_naaped.Cells(i + 2, "I").Value = "=DataSet_" & i + 1 & "!Z34"
        'Current
        sh_naaped.Cells(i + 2, "J").Value = "=DataSet_" & i + 1 & "!U1"
    Next i

    i_av = i + 1

    'Averages
    For ii = 0 To copy_cols - 1
        sh_naaped.Range("L2").Offset(0, ii).Value = "=AVERAGE(C2:" & sh_naaped.Cells(i_av, "C").Address(False, False) & ")"
        sh_naaped.Range("L3").Offset(0, ii).Value = "=STDEV(C2:" & sh_naaped.Cells(i_av, "C").Address(False, False) & ")"
        sh_naaped.Range(sh_naaped.Cells(2, "L"), sh_naaped.Cells(3, "S")).FillRight
    Next ii
    sh_data.Select
End Sub




