***VBA_MISCELLANEOUS***

A collection of VBA code I made to serve multilpe automation task during my work. 

**1. Xlookup/vlookup equivalent on VBA**
 
**Context**: We want to do a lookup using data from col a,b,c and perform on col d (respectively), to see which email has already existed in our database, the found values are returned in col E. A native excel formula would look like: 

```
E2=XLOOKUP(A2,D:D,D:D,XLOOKUP(B2,D:D,D:D,XLOOKUP(C2,D:D,D:D,"",0,1)))
```
**Explanation**: lookup A2, in array D, if found return the found values in col E, If not proceed to lookup in B2 and C2. In this case the found emails are in A2, B3, C4.

For easier illustration, please refer to the screenshot (ALL email data is fake and only use for the purpose of illustration)
<p align="center">
<img src="https://user-images.githubusercontent.com/125301325/228885683-1d2d8039-a054-4b81-80ea-057669537662.png"
width="650">
</p>

**Question**: Why don't we just use the xlookup built-in function which is similar to the native formula and actually easier to do?
 
**Answer**: the native xlookup function is very useful in Excel, but in VBA, maybe it's just me but I find it hard to loop the code, plus typing the code is far too consuming, especially if you are handling a bulk or documents. So I used the alternative Find.function in this case. Of course you can always perform normal xlookup.function in VBA, instruction is [here](https://www.automateexcel.com/vba/vlookup-xlookup/#:~:text=The%20VLOOKUP%20and%20XLOOKUP%20functions%20in%20Excel%20are%20extremely%20useful,be%20used%20in%20VBA%20Coding.)

Let's begin!

**Step 1: Declaring variables** 
First we need to to declare the  variables and setup the macro, data type, I would assume that you have already known about this. If not you can read more [here](https://learn.microsoft.com/en-us/power-automate/desktop-flows/variable-data-types.) 

```
Sub FindDuplicate()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim l As Long
    Dim j As Long, 
    Dim searchValue As Variant
    Dim searchRange As Range
    Dim foundCell As Range
    Dim lastRowA As Long
    Dim lastRowB As Long
    Dim lastRowC As Long
    Dim highestLastRow As Long
```
**Explanation**: j is the row in column A being looped, x is column E, l is for row in column D

**Step 2: Initializing variables**
```
Set ws = ThisWorkbook.Worksheets("Sheet1") 'change this to your desired worksheet name'
lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row 'reminder: D col will be the col that we perform the lookup on'
lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
lastRowC = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
```
  
**Step 3: Setting the searchvalue, searchrange and function**

```
    For l = 1 To lastRow 'counting the row in col D till last row which already set up in step 2
    For j = 0 To 2 
    searchValue = ws.Cells(l, j + 1).Value
        
        Set searchRange = ws.Range("D2:D6") 'change to fit your desired search range'
        
        Set foundCell = searchRange.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole)
```
**Explanation**: 
l is the row in col D which is the row to be peformed lookup on, j is the outerloop for iteration between columns, 
setting searchValues = Cells function(rowindex, column index) in this case it's ws.Cells(l, j+1) (remember that j+1 is the location of the column)
* On the first iteration of j loop, column position will be 0 + 1 =1, which is column A (1st position)
* On second iteration of j loop, column position will be 1 + 1 = 2, which is column B (2nd position)
* On third  iteration of j loop, column position will be 1 + 2 = 2, which is column C (3rd position)

Thus we have completed the loop for 3 columns, the code will use data from A,B,C and lookup in column D

_note_: if you want to loop through more columsn you can change it for the j outer loop (for example: j =  0 to 3 will loop the column 4 times in the fourth iteration, column position will be 3 + 1 =4 (4th postion) which equals column D)

Set setting range using range function: ws.Range ("D2:D6") 

Setting foundcell using the built-in Find fuction 
Original formula is as follows: 
```
expression.Find (What, After, LookIn, LookAt, SearchOrder (optional), SearchDirection (optional), MatchCase (optional), MatchByte (optional), SearchFormat(optional)
```
the expression precedes .Find is supposed to be a range object (in this case it's our searchRange), data to searchfor is searchValue which already been defined. Read more about the Find function [here](https://learn.microsoft.com/en-us/office/vba/api/excel.range.find). 

**Step 4: Setting the If condition if values are found**
```
        ' If a match is found, enter the value in column E that matches the row it was found'
        If Not foundCell Is Nothing Then
            ws.Cells(l, "E").Value = foundCell
        End If
    Next j
    Next l
End sub
```
**Explanation**: 

using the function .Cells, we set up the column to return the found values on (which is column 'E'). We again use the function .Cells(rowindex, columnindex) to pinpoint the cells that we want to return the values in. In this case the rowindex is 'l' which we've already initialized in step 3. 
The next j and next l is for the formula to loop through the next columns/rows.
Read more about the If. function [here](https://www.automateexcel.com/vba/else-if-statement)


**Complete code**

```
Sub FindDuplicate()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim l As Long 'you can dim whatever alias letter here'
    Dim j As Long 'you can dim whatever alias letter here'
    

    Dim searchValue As Variant
    Dim searchRange As Range
    Dim foundCell As Range
    Dim lastRowA As Long
    Dim lastRowB As Long
    Dim lastRowC As Long
    Dim highestLastRow As Long
    
    ' Initialize variables
    Set ws = ThisWorkbook.Worksheets("Sheet1") 'change this to fit your worksheet name'
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastRowC = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
  
    
    ' Loop through the values in column B
    
    For l = 1 To lastRow
    For j = 0 To 2
    searchValue = ws.Cells(l, j + 1).Value 
        
        Set searchRange = ws.Range("D2:D6") 'change the range to fit your required searchRange'
        
        Set foundCell = searchRange.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' If a match is found, enter the value in column E that matches the row it was found'
        If Not foundCell Is Nothing Then
            ws.Cells(l, "E").Value = foundCell 'change the return column to your desired location'
        End If
    Next j
    Next l
    
End Sub
```
It's not the most sophisticated lookup ever, but it does the job, if you find it too hard to understand (I would) you can just adjust certain variables to your desired values (e.g., workhsheet, range).
