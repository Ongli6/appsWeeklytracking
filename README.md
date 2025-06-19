# appsWeeklytracking
VBA code for a self generating Excel sheet that tracks Applications
Coded in VBA. 

Spreadsheet contained 4 hidden sheets containing pivot tables of data imported from powerbi, 2 Weekly reports (16-18 and 19+) aswell as 6 campus breakdown sheets (16-18 and 19+ for each campus).
Campus totals are tracked in 8 block, Faculty breakdowns are tracked in 6 Blocks. So seperate buttons/logic used to do both. Seperated as it was easier to test just one button as opposed to running the whole code from the top each time.

Logic basically copies the template found in the first 6/8 columns, copies in formula extracted from pivot table on hidden sheets and then copies and pastes values to consolidate figures.

Some extra faffing around copying specific merge cell formatting included in both 6 and 8 with the latter being more complex as extra two columns were being difficult during testing (probably as its a fatty merged cell)
