Attribute VB_Name = "Module1"

Sub FormatTrialBalance()
Attribute FormatTrialBalance.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim tbRange As Range
    
    ' Define the trial balance range starting at A1
    Set tbRange = Range("A1").CurrentRegion
    
    ' Format headers (first row)
    With tbRange.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220) ' light gray
        .HorizontalAlignment = xlCenter
    End With
    
    ' Format numbers in Debit and Credit columns (C and D)
    With tbRange.Columns("C:D")
        .NumberFormat = "#,##0.00"
    End With
    
    ' Apply borders to the whole table
    With tbRange.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    ' Auto-fit columns
    tbRange.Columns.EntireColumn.AutoFit

End Sub
