Sub Test()
'
' Test 宏
'
' 快捷键: Ctrl+g
'

    Dim Total_i, Data_i, Data_j, i, Cur_file As Integer
    Dim temp_i As Integer
    
    Dim Cur_num
    Dim temp2
    Dim Spot_num
    Dim Data(16, 80)
    Dim Cur_data(16, 5000)
    Dim MSD(16, 1)
    Dim Max_x
    Dim Cur_MSD
    Dim temp
    Dim b
    
    Dim File_num
    Dim Saving_Path
    
    i = 0
    b = 0
    
    Saving_Path = InputBox("Enter files location:", , "C:\Users\gaoyuan\Desktop\MSD\")
    File_num = Int(InputBox("Enter Files Number:"))
    
    For Cur_file = 1 To File_num
    
    '复制为数据
    
    ChDir Saving_Path + CStr(Cur_file)
    Workbooks.OpenXML Filename:= _
        Saving_Path + CStr(Cur_file) + "\Experiment-2170_Tracks.xml", LoadOption:= _
        xlXmlLoadImportToList
    Columns("A:K").Select
    Selection.Copy
    Windows("Test.xlsm").Activate
    Sheets("Sheet2").Select
    Cells(1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    '找点
    Total_num = Range("A65536").End(xlUp).Row
    Max_x = Application.WorksheetFunction.Max(Range(Cells(2, 9), Cells(Total_num, 9)))
    Spot_num = 0
    Cur_num = -1
    
    For Total_i = 2 To Total_num
        If (Cur_num <> Sheet2.Cells(Total_i, 7) And Sheet2.Cells(Total_i, 7) > 17) Then
            
            Cur_num = Sheet2.Cells(Total_i, 7)
            
            Dim a
            
            'a = Application.WorksheetFunction.Var(Range(Cells(Total_i, 9), Cells(Total_i + Cur_num - 1, 9)))
                
            'Sheet1.Cells(Spot_num + 1, Cur_file + 2) = Application.WorksheetFunction.Correl(Sheet2.Range(Cells(Total_i, 8), Cells(Total_i + Cur_num - 1, 8)), Sheet2.Range(Cells(Total_i, 9), Cells(Total_i + Cur_num - 1, 9)))
            
            a = Application.WorksheetFunction.Correl(Sheet2.Range(Cells(Total_i, 8), Cells(Total_i + Cur_num - 1, 8)), Sheet2.Range(Cells(Total_i, 9), Cells(Total_i + Cur_num - 1, 9)))
            
            If Abs(a) > 0.7 Then
                
                If Sheet2.Cells(Total_i, 7) < 20 Then
                        
                        'MsgBox (Sheet2.Cells(Total_i, 7))
                        
                        Cur_MSD = 0
                        
                        For i = 0 To 16
                
                
                        '累加
                        Cur_MSD = Cur_MSD + (Cells(Total_i + i, 9) - Cells(Total_i, 9)) ^ 2 + (Cells(Total_i + i, 10) - Cells(Total_i, 10)) ^ 2
                        Cur_data(i, Spot_num) = Cur_MSD / 15
                
                        'Cur_data(i, Spot_num) = (Cells(Total_i + i, 9) - Cells(Total_i, 9)) ^ 2 + (Cells(Total_i + i, 10) - Cells(Total_i, 10)) ^ 2
                
                
                
                        Next
                
                        Spot_num = Spot_num + 1
                        'MsgBox (Spot_num)
                
                Else
                        
                        'MsgBox ("执行！" + CStr(Spot_num))
                        
                        temp = Sheet2.Cells(Total_i, 7)
                        temp2 = Total_i + temp - 18
                        'temp_i = Total_i
                        
                        For temp_i = Total_i To temp2
                            
                            Dim c
                            
                            c = Application.WorksheetFunction.Correl(Sheet2.Range(Cells(temp_i, 8), Cells(temp_i + 15, 8)), Sheet2.Range(Cells(temp_i, 9), Cells(temp_i + 15, 9)))
                            
                            If Abs(c) > 0.7 Then
                            
                                'MsgBox ("执行！" + CStr(Spot_num))
                                Cur_MSD = 0
                            
                                For i = 0 To 16
                
                
                                    '累加
                                    Cur_MSD = Cur_MSD + (Cells(temp_i + i, 9) - Cells(temp_i, 9)) ^ 2 + (Cells(temp_i + i, 10) - Cells(temp_i, 10)) ^ 2
                                    Cur_data(i, Spot_num) = Cur_MSD / 15
                
                                    'Cur_data(i, Spot_num) = (Cells(Total_i + i, 9) - Cells(Total_i, 9)) ^ 2 + (Cells(Total_i + i, 10) - Cells(Total_i, 10)) ^ 2
                
                
                                Next
                
                                Spot_num = Spot_num + 1
                                'MsgBox ("执行！" + CStr(Spot_num))
                            
                            End If
                            
                            'Sheet1.Cells(Spot_num + 1, Cur_file + 2) = c
                            
                       Next
                        
                End If
                
                'MsgBox ("执行！" + CStr(Spot_num))
            
            End If
            
        End If
        
    Next
    
    
    Columns("A:K").Select
    Selection.ClearContents
    
    temp2 = 0
    
    For Data_i = 0 To 16
        
        temp = 0
        
        For Data_j = 0 To Spot_num - 1
        
            temp = temp + Cur_data(Data_i, Data_j)
            
        Next
        
        If Spot_num <> 0 Then
        
            Data(Data_i, Cur_file - 1) = temp / Spot_num
        
        Else
            
            Data(Data_i, Cur_file - 1) = 0
            temp2 = 1
            
        End If
        
    Next
    
    Sheets("Sheet1").Select
    MsgBox (Spot_num)
    
    If temp2 Then
        
        b = b + 1
        
    End If
        
    
    'Range("b2").Resize(UBound(Cur_data, 1), UBound(Cur_data, 2)) = Cur_data
    Erase Cur_data
    
    Next
    
    For Data_i = 0 To 16
        
        temp = 0
        
        For Data_j = 0 To File_num - 1
        
            temp = temp + Data(Data_i, Data_j)
            
        Next
        
        MSD(Data_i, 0) = temp / (File_num - b)
    
    Next
    
    Range("b2").Resize(UBound(MSD, 1), 1) = MSD
    'Range("b2").Resize(UBound(Cur_data, 1), UBound(Cur_data, 2)) = Cur_data
    
    
    '画图
    Range("A1:B17").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlXYScatterSmooth
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$1:$B$17")
    
End Sub
