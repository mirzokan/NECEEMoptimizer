Option Explicit

'Global variable declaration
    'Sheets
Public sh_data As Worksheet
Public sh_settings As Worksheet
Public sh_antemp As Worksheet
Public sh_calcval As Worksheet
Public sh_naaped As Worksheet
Public sh_report As Worksheet
Public sh_instructions As Worksheet
Public sh_dataex As Worksheet

Sub Initialize()
 'Procedure re-initializes the Global Variables

    'Initialize sheets
    Set sh_data = ThisWorkbook.Sheets("Data Input")
    Set sh_settings = ThisWorkbook.Sheets("hidden_settings")
    Set sh_antemp = ThisWorkbook.Sheets("hidden_analysis_template")
    Set sh_calcval = ThisWorkbook.Sheets("hidden_calcval")
    Set sh_naaped = ThisWorkbook.Sheets("hidden_naaped")
    Set sh_report = ThisWorkbook.Sheets("Optimization Report")
    Set sh_instructions = ThisWorkbook.Sheets("Instructions")
    Set sh_dataex = ThisWorkbook.Sheets("Example_Data")
End Sub

Sub Reset()
    Dim i As Long, arr_el As Long, ds_num  As Long
    Dim ds_names() As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Call Initialize

    'Delete data entries from NAAPed table
    sh_naaped.Range("A2:S" & sh_naaped.Range("B2").End(xlDown).Row).ClearContents

    'Reset optimization overrides
    sh_calcval.Range("opt_overrides").Value = 0
    sh_calcval.Range("iter_rec").Value = 0

    'Reset talker cell
    sh_report.Range("talker:F31").ClearContents

    'Delete data sheets
    ds_num = 0
    For i = 1 To Worksheets.Count
        If Left(Sheets(i).Name, 7) = "DataSet" Then
        ds_num = ds_num + 1
        End If
    Next i

    If ds_num >= 1 Then
        ReDim ds_names(ds_num - 1)
        arr_el = 0
        For i = 1 To Worksheets.Count
            If Left(Sheets(i).Name, 7) = "DataSet" Then
                ds_names(arr_el) = Sheets(i).Name
                arr_el = arr_el + 1
            End If
        Next i

        For i = 0 To UBound(ds_names)
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets(ds_names(i)).Delete
            Application.DisplayAlerts = True
        Next i
    End If
  
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub

Sub Button_Reset()
    Call Reset
    MsgBox ("The sheet has been reset!")
End Sub

Sub Button_Optimize()
    Dim file_list() As Variant
    Dim file_path As Range
    Dim file_count As Long

    ' Start fresh
    Call Reset

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' Manual or automatic data treatment
    If sh_data.Range("pref_data") = sh_settings.Range("A6") Then
        sh_antemp.Visible = xlSheetVisible
        Set file_path = Range("datapath")
        Call validate_filepath(file_path)
        file_list = list_files(file_path.Value)
        file_count = UBound(file_list) + 1
        Call ImportWorksheets(file_path, file_list)
        Call print_filelist(file_list)

        Application.Calculate

        'Cleanup
        Erase file_list
    End If
    
    Call optimize
    ' Call check_fails
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    sh_antemp.Visible = xlSheetHidden
    sh_report.Activate
    MsgBox ("Analysis Complete!")
End Sub

Sub Button_Recalculate()
    Dim iCount As Integer, i As Integer
    Dim sText As String
    Dim lNum As String
    Dim sheet_num As Long
   
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ActiveSheet.Calculate

   'Identify sheet number
    sText = ActiveSheet.Name
    For iCount = Len(sText) To 1 Step -1
        If IsNumeric(Mid(sText, iCount, 1)) Then
            i = i + 1
            lNum = Mid(sText, iCount, 1) & lNum
        End If
         
        If i = 1 Then lNum = CInt(Mid(lNum, 1, 1))
    Next iCount
    
    sheet_num = CLng(lNum)
    clear_template (sheet_num)
    adjust_template (sheet_num)

    ActiveSheet.Calculate
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    MsgBox ("Recalculation complete!")
End Sub

Sub toggle_instructions()
    If sh_instructions.Visible = xlSheetVisible Then
        sh_instructions.Visible = xlSheetHidden
        sh_data.Activate
    Else
        sh_instructions.Visible = xlSheetVisible
        sh_instructions.Activate
    End If
End Sub

Sub toggle_dataex()
    If sh_dataex.Visible = xlSheetVisible Then
        sh_dataex.Visible = xlSheetHidden
        sh_data.Activate
    Else
        sh_dataex.Visible = xlSheetVisible
        sh_dataex.Activate
    End If
End Sub

