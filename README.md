# VBA-Challenge

2nd challenge of the UTSA Data Analytics Bootcamp
Utilizes Microsoft Excel VBA coding

Below are code used that was researched to meet the goals of the challenge:

From https://www.excelsirji.com/vba-code-to-change-cell-color/,

'Change cell color to green
    Sheet1.Range("C4").Interior.Color = RGB(0, 255, 0)
    
'Change cell color to red
    Sheet1.Range("C5").Interior.Color = RGB(255, 0, 0)

From https://www.thespreadsheetguru.com/vba-determine-if-cell-has-no-fill-color/,

ActiveCell.Interior.ColorIndex = xlNone

The above code was used for formatting cells depending on the yearly change.

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'Countdown of rows of Whole Dataset

This code was used to find the last row of the given dataset. First knew about this code via Ashelyn Allred, fellow Bootcamp colleague.
