# VBA-challenge
Module 2 Challenge Code contains all code that I used to run my script. I have 2 Subs: Stock_Outputs and MaxMinVals. 
The first sub: Stock_Outputs contains the columns created for the output table and the loops/conditionals to correctly summarize the outputs required for the Module 2 Challenge. 
The second sub: MaxMinVals is for the last part of the challenge. Outputing the largest/smallest pecent increases and decreases, and the largest total stock volume.
References: https://www.mrexcel.com/board/threads/max-min-vba.132404/
I used this website above to learn how to apply excel formulas in VBA. I wanted to apply the max and min formulas to return the highest/lowest percent increase/decrease in the output table. This is seen in lines 53-55 of my code:
    max_p = WorksheetFunction.Max(ws.Range("K:K"))
    min_p = WorksheetFunction.Min(ws.Range("K:K"))
    max_vol = WorksheetFunction.Max(ws.Range("L:L"))
Once I learned how to apply Excel functions in VBA through the website, I used that knowledge to apply a lookup fucntion to return the ticker symbol by looking up the max/min outputs in lubes 61-63 of my code:
    ws.Cells(2, "O").Value = WorksheetFunction.XLookup(ws.Cells(2, "P"), ws.Range("K:K"), ws.Range("I:I"))
    ws.Cells(3, "O").Value = WorksheetFunction.XLookup(ws.Cells(3, "P"), ws.Range("K:K"), ws.Range("I:I")) 
    ws.Cells(4, "O").Value = WorksheetFunction.XLookup(ws.Cells(4, "P"), ws.Range("L:L"), ws.Range("I:I"))
I also used the record Macro tool in excel to learn how to apply certain formats to cells/ranges which I used in my code. This is seen in lines 24,31,32,59,60: 
    ws.Cells(out_range, "J").NumberFormat = "0.00"
    ws.Cells(out_range, "K").Style = "Percent"
    ws.Cells(out_range, "K").NumberFormat = "0.00%"
    ws.Range("P2:P3").Style = "Percent"
    ws.Range("P2:P3").NumberFormat = "0.00%"
    
