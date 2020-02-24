Dim copy_from_prefix As String                      'data tab name prefix
Dim copy_from_suffix As Variant                     'array of data tab name suffix
Dim copy_to As String                               'destination tab name
Dim text_file_name As String                        'file name for compiled data text file
Dim text_file_name_and_path As String               'concat working file path and text file name
Dim last_row As Integer                             'last row with data on destination tab
Dim next_row As Integer                             'offset row count for new data on destination tab
Dim copy_to_dim_col_offset As Integer               'offset column count for dimension columns on destination tab
Dim copy_to_metric_col_offset As Integer            'offset column count for metric columns on destination tab
Dim index_num As Integer                            'used to iterate over array of data tabs
Dim start_week As Integer                           'data start week number
Dim end_week As Integer                             'data end week number
Dim week_num As Integer                             'iterates over each week between start and end week; invokes calculation in working file for each week
Dim data_start_row As Integer                       'row number from which data starts on data tabs
Dim data_end_row As Integer                         'row number on which data ends on data tabs
Dim dim_start_col As Integer                        'column number from which dimension column starts on data tabs
Dim dim_end_col As Integer                          'column number on which dimension column ends on data tabs
Dim pkg_start_col As Integer                        'column number from which package metric column starts on data tabs
Dim pkg_end_col As Integer                          'column number on which package metric column ends on data tabs
Dim segA_start_col As Integer                       'column number from which seg auto metric column starts on data tabs
Dim segA_end_col As Integer                         'column number on which seg auto metric column ends on data tabs
Dim segP_start_col As Integer                       'column number from which seg prop metric column starts on data tabs
Dim segP_end_col As Integer                         'column number on which seg prop metric column ends on data tabs
Dim current_wb As Object                            'working file workbook containing macro
Dim temp_wb As Object                               'temporary workbook used to copy compiled data and save as flat file
Dim target_file As Object                           'excel workbook containing target values

'calculate weekly target values in Working file based on target values in Target workbook
'copy calculated weekly target data tabs to destination tab
'after iterating process for all weeks, save compiled tab as flat file to be uploaded to Hadoop
Sub Compile_Trgt()

    Application.ScreenUpdating = False              'stop screen flicker
    
    'assign values from excel worksheet using named cell ranges
    copy_from_prefix = Range("rs_copy_from_prefix").Value
    copy_from_suffix = Array(Range("rs_copy_from_suffix_1").Value, Range("rs_copy_from_suffix_2").Value, Range("rs_copy_from_suffix_3").Value)
    copy_to = Range("rs_copy_to").Value
    data_start_row = Range("Data_Start_Row").Value
    data_end_row = Range("Data_End_Row").Value
    dim_start_col = Range("Dim_Start_Col").Value
    dim_end_col = Range("Dim_End_Col").Value
    pkg_start_col = Range("Pkg_Start_Col").Value
    pkg_end_col = Range("Pkg_End_Col").Value
    segA_start_col = Range("SegA_Start_Col").Value
    segA_end_col = Range("SegA_End_Col").Value
    segP_start_col = Range("SegP_Start_Col").Value
    segP_end_col = Range("SegP_End_Col").Value
    copy_to_dim_col_offset = Range("dim_col_offset").Value
    copy_to_metric_col_offset = Range("metric_col_offset").Value
    start_week = Range("rs_start_week").Value
    end_week = Range("rs_end_week").Value
    text_file_name = "Risk_State_Target"
    next_row = 1                                    'initial row number for new data
    
    'set current_wb to Working file
    Set current_wb = ActiveWorkbook
    
    'delete preexisting data from destination tab to copy over new values
    'if Cell A2 is blank then skip delete step
    ThisWorkbook.Worksheets(copy_to).Activate
    If Range("A2").Value = "" Then
        last_row = 1
    Else
        Cells(1, 1).Select
        Selection.End(xlDown).Select
        last_row = ActiveCell.Row                   'used in delete statement below
    End If
    'delete rows from destination (copy_to) tab
    If last_row > 1 Then                            'to skip deleting headers
        Application.StatusBar = "Deleting Existing Destination Info..."
        ActiveSheet.Range(Cells(1, 1).Offset(next_row, 0), Cells(1, 1).Offset(last_row, 0)).EntireRow.Delete
    Else 
        Cells(1, 1).Select
    End If

    'open workbook containing target values for efficient formula calculation in Working file
    Workbooks.Open Filename:=Range("rs_targetfile").Value, UpdateLinks:=0, ReadOnly:=True
    Set target_file = ActiveWorkbook                '
    
    '
    current_wb.Activate                             'activate Working file
    'access 'copy_from_suffix' array's elements
    For index_num = LBound(copy_from_suffix) To UBound(copy_from_suffix)            
        week_num = start_week
        
        'QQ is not split by LOB and therefore it only needs to copy-paste one section
        If copy_from_suffix(index_num) <> "QQ" Then     

            'if suffix value is Issued or FQ then below Do Loop runs
            'uses Offsets to find appropriate copy and paste range
            Do
                Application.StatusBar = "Copying from " & copy_from_prefix & copy_from_suffix(index_num) & " week " & week_num
                
                'sets value for week number in source tab and calculates the worksheet
                Range("rs_week").Value = week_num
                ThisWorkbook.Worksheets("MacroForm").Calculate
                ThisWorkbook.Worksheets(copy_from_prefix & copy_from_suffix(index_num)).Activate
                ThisWorkbook.Worksheets(copy_from_prefix & copy_from_suffix(index_num)).Calculate
        
                'copy dimension fields from source tab and paste in destination tab
                ThisWorkbook.Worksheets(copy_from_prefix & copy_from_suffix(index_num)).Activate
                Range(Cells(1, 1).Offset(data_start_row - 1, dim_start_col - 1), Cells(1, 1).Offset(data_end_row - 1, dim_end_col - 1)).Copy
                ThisWorkbook.Worksheets(copy_to).Activate
                ActiveSheet.Cells(1, 1).Offset(next_row, copy_to_dim_col_offset).PasteSpecial xlPasteValues
        
                'copy package metrics fields from source tab and paste in destination tab
                ThisWorkbook.Worksheets(copy_from_prefix & copy_from_suffix(index_num)).Activate
                Range(Cells(1, 1).Offset(data_start_row - 1, pkg_start_col - 1), Cells(1, 1).Offset(data_end_row - 1, pkg_end_col - 1)).Copy
                ThisWorkbook.Worksheets(copy_to).Activate
                ActiveSheet.Cells(1, 1).Offset(next_row, copy_to_metric_col_offset).PasteSpecial xlPasteValues
        
                'row after above data pasted on destination tab
                next_row = next_row + (data_end_row - data_start_row + 1)
        
                'copy dimension fields from source tab and paste in destination tab
                ThisWorkbook.Worksheets(copy_from_prefix & copy_from_suffix(index_num)).Activate
                Range(Cells(1, 1).Offset(data_start_row - 1, dim_start_col - 1), Cells(1, 1).Offset(data_end_row - 1, dim_end_col - 1)).Copy
                ThisWorkbook.Worksheets(copy_to).Activate
                ActiveSheet.Cells(1, 1).Offset(next_row, copy_to_dim_col_offset).PasteSpecial xlPasteValues
        
                'copy segment auto metrics fields from source tab and paste in destination tab
                ThisWorkbook.Worksheets(copy_from_prefix & copy_from_suffix(index_num)).Activate
                Range(Cells(1, 1).Offset(data_start_row - 1, segA_start_col - 1), Cells(1, 1).Offset(data_end_row - 1, segA_end_col - 1)).Copy
                ThisWorkbook.Worksheets(copy_to).Activate
                ActiveSheet.Cells(1, 1).Offset(next_row, copy_to_metric_col_offset).PasteSpecial xlPasteValues
        
                'row after above data pasted on destination tab
                next_row = next_row + (data_end_row - data_start_row + 1)
        
                'copy dimension fields from source tab and paste in destination tab
                ThisWorkbook.Worksheets(copy_from_prefix & copy_from_suffix(index_num)).Activate
                Range(Cells(1, 1).Offset(data_start_row - 1, dim_start_col - 1), Cells(1, 1).Offset(data_end_row - 1, dim_end_col - 1)).Copy
                ThisWorkbook.Worksheets(copy_to).Activate
                ActiveSheet.Cells(1, 1).Offset(next_row, copy_to_dim_col_offset).PasteSpecial xlPasteValues
        
                'copy segment property metrics fields from source tab and paste in destination tab
                ThisWorkbook.Worksheets(copy_from_prefix & copy_from_suffix(index_num)).Activate
                Range(Cells(1, 1).Offset(data_start_row - 1, segP_start_col - 1), Cells(1, 1).Offset(data_end_row - 1, segP_end_col - 1)).Copy
                ThisWorkbook.Worksheets(copy_to).Activate
                ActiveSheet.Cells(1, 1).Offset(next_row, copy_to_metric_col_offset).PasteSpecial xlPasteValues
       
                'row after above data pasted on destination tab
                next_row = next_row + (data_end_row - data_start_row + 1)
        
                'next week
                week_num = week_num + 1
        
            Loop Until week_num > end_week
        
        Else
            week_num = start_week
            
            'if suffix value is QQ then below Do Loop runs
            'uses Offsets to find appropriate copy and paste range
            Do
                Application.StatusBar = "Copying from " & copy_from_prefix & copy_from_suffix(index_num) & " week " & week_num
                
                'sets value for week number in source tab and calculates the worksheet
                Range("rs_week").Value = week_num
                ThisWorkbook.Worksheets("MacroForm").Calculate
                ThisWorkbook.Worksheets(copy_from_prefix & copy_from_suffix(index_num)).Activate
                ThisWorkbook.Worksheets(copy_from_prefix & copy_from_suffix(index_num)).Calculate
        
                'copy dimension fields from source tab and paste in destination tab
                ThisWorkbook.Worksheets(copy_from_prefix & copy_from_suffix(index_num)).Activate
                Range(Cells(1, 1).Offset(data_start_row - 1, dim_start_col - 1), Cells(1, 1).Offset(data_end_row - 1, dim_end_col - 1)).Copy
                ThisWorkbook.Worksheets(copy_to).Activate
                ActiveSheet.Cells(1, 1).Offset(next_row, copy_to_dim_col_offset).PasteSpecial xlPasteValues
        
                'copy total metrics fields for QQ
                ThisWorkbook.Worksheets(copy_from_prefix & copy_from_suffix(index_num)).Activate
                Range(Cells(1, 1).Offset(data_start_row - 1, pkg_start_col - 1), Cells(1, 1).Offset(data_end_row - 1, pkg_end_col - 1)).Copy
                ThisWorkbook.Worksheets(copy_to).Activate
                ActiveSheet.Cells(1, 1).Offset(next_row, copy_to_metric_col_offset).PasteSpecial xlPasteValues
        
                'row after above data pasted on destination tab
                next_row = next_row + (data_end_row - data_start_row + 1)
                
                'next week
                week_num = week_num + 1
        
            Loop Until week_num > end_week
        End If    
    Next index_num
    
    Application.CutCopyMode = False                 'to clear the clipboard
    target_file.Close savechanges:=False
    Application.StatusBar = "Creating Text File"
    
    'create a text file for "copy to" tab
    current_wb.Activate
    text_file_name_and_path = current_wb.Path & "\" & text_file_name

    current_wb.Worksheets(copy_to).Activate
    current_wb.ActiveSheet.UsedRange.Copy
    
    Set temp_wb = Application.Workbooks.Add         
    temp_wb.Sheets(1).Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False             
    
    Application.DisplayAlerts = False               'suppress prompts and alert messages while a macro is running
    
    'xlText will save as Tab Deliminated text
    temp_wb.SaveAs Filename:=text_file_name_and_path, FileFormat:=xlText, ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges
    temp_wb.Close                                   'close temp workbook
    
    current_wb.Save

    Application.StatusBar = "Done!!"
    MsgBox ("New " & text_file_name & ".txt file saved to the folder...")
End Sub
