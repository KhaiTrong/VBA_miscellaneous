##VBA_miscellaneous##

A collection of VBA code I made to serve multilpe automation task in my work

**1. Xlookup/vlookup equivalent on VBA**
 
**Context**: We want to do a lookup using data from col a,b,c and perform on col d (respectively), to see which email has already existed in our database, the found values are returned in col E. A native excel formula would look like: 

```
E2=XLOOKUP(A2,D:D,D:D,XLOOKUP(B2,D:D,D:D,XLOOKUP(C2,D:D,D:D,"",0,1)))
```
**Explaination**: lookup A2, in array D, if found return the found values in col E, If not proceed to lookup in B2 and C2. In this case the found emails are in A2, B3, C4.

For easier illustration, please refer to the screenshot (ALL email data is fake and only use for the purpose of illustration)
<p align="center">
<img src="https://user-images.githubusercontent.com/125301325/228873083-337334f0-32c9-4a89-8f43-8ec435ad72a4.png"
width="650">
</p>

**Question**: Why don't we just use the xlookup built-in function which is similar to the native formula and actually easier to do?
 
**Answer**: the native xlookup function is very useful in Excel, but in VBA, maybe it's just me but I find it hard to loop the code, plus typing the code is far too consuming, especially if you are handling a bulk or documents. So I used the alternative Find.function in this case. Of course you can always perform normal xlookup.function in VBA, instruction is [here](https://www.automateexcel.com/vba/vlookup-xlookup/#:~:text=The%20VLOOKUP%20and%20XLOOKUP%20functions%20in%20Excel%20are%20extremely%20useful,be%20used%20in%20VBA%20Coding.)

Let's begin!

**Declaring variables** 
First we need to to declare the  variables and setup the macro, data type, I would assume that you have already known about this. If not you can read more [here](https://learn.microsoft.com/en-us/power-automate/desktop-flows/variable-data-types.) 

```
Sub FindDuplicate()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim l As Long
    Dim j As Long, x As Long: x = 2
    Dim searchValue As Variant
    Dim searchRange As Range
    Dim foundCell As Range
    Dim lastRowA As Long
    Dim lastRowB As Long
    Dim lastRowC As Long
    Dim highestLastRow As Long
```
**Explaination**: j is the row in column A being looped, x is column E, l is for

**Initializing variables**
```
Set ws = ThisWorkbook.Worksheets("Sheet1") 'change this to your desired worksheet name'
lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
lastRowC = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
```
  
**Explaination**: setting which worksheet we wish to work on & counting the 
