<div align="center">

## Column Print


</div>

### Description

Print in columns to a printer or adapt this code to print to a form. First character alignes in a column for each row. No "Space(##)" and no grid.

The project I am working on needed to be able to print an unforseen number of single words enterd by the user, whether 5, 25 or 100 and it should be neat and consistant. This code is a cut down version of my sub to show the principle and can be easily adapted. My thanks go to Harvest R for a tutorial(80647232000.zip) I downloaded in Aug/00

which I highly recomend. I hope this code can be used as a learning aid.
 
### More Info
 
In the code ### would be a veriable in another part of your program. I also have a limit on the Text Box where words are typed in, to make sure they will fit into the 48mm column width I have set. This can easily be changed to suit your own needs, 2 columns 3,5,12 whatever.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mark Hunter](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-hunter.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mark-hunter-column-print__1-11447/archive/master.zip)





### Source Code

```
Private Sub mnuFilePrint_Click()
Dim TodaysDate AS Variant
Dim HorizontalMargin As Single
Dim VerticalMargin As Single
Dim BeginPage As Single, EndPage As Single
Dim NumCopies As Single
Dim SheetStyleText As String 'WorkSheet
Dim SheetStyleTextWidth As Single
Dim JobNumberText As String 'JobNo ###
Dim JobNumberTextWidth As Single
Dim CompanyNameText As String '###
Dim CompanyNameTextWidth As Single
Dim JobDescriptionText As String
Dim JobDescriptionTextWidth As Single
Dim JobFontText As String '###
Dim JobFontTextWidth As Single
Dim JobFontSizeText As String '###
Dim JobFontSizeTextWidth As Single
Dim t As Integer  ' copies
Dim f As Integer  ' counter for anystring()
Dim k As Integer	'counter for column's
Dim Col(0 To 3), NR	 '4 column's and next row
 CommonDialog1.CancelError = True
 On Error GoTo ErrHandler
	' Display the Print dialog box
 CommonDialog1.ShowPrinter
 ' Get user-selected values from the dialog box
Printer.ScaleMode = 6   'millimeters
HorizontalMargin = CommonDialog1.PrinterDefault
VerticalMargin = CommonDialog1.PrinterDefault
BeginPage = CommonDialog1.FromPage
EndPage = CommonDialog1.ToPage
NumCopies = CommonDialog1.Copies
For t = 1 To NumCopies
Next t
HorizontalMargin = 10 + HorizontalMargin
VerticalMargin = 5 + VerticalMargin
Printer.FontName = "Arial"
Printer.FontSize = 12
Printer.FontBold = True
Printer.FontItalic = False
Printer.FontUnderline = False
Printer.FontStrikethru = False
Printer.ForeColor = RGB(0, 0, 0)
TodaysDate = Format(Date, "Long Date")
Printer.Print "Header Name"; Space(110); 'initialize the printer
Printer.Print TodaysDate
Printer.FontName = "Arial"
Printer.FontSize = 16
Printer.FontBold = True
Printer.FontItalic = False
Printer.FontUnderline = False
Printer.FontStrikethru = False
Printer.ForeColor = RGB(0, 0, 0)
CompanyNameText = "XYZ Company & Co" 'user name###
CompanyNameTextWidth = Printer.TextWidth(CompanyNameText)
Printer.CurrentX = (210-CompanyNameTextWidth) / 4
Printer.CurrentY = VerticalMargin + 15
Printer.Print CompanyNameText
SheetStyleText1 = "Work Sheet"
SheetStyleTextWidth1 = Printer.TextWidth(SheetStyleText1)
Printer.CurrentX = (210-SheetStyleTextWidth1)/1.5
Printer.CurrentY = VerticalMargin + 15
Printer.Print SheetStyleText1
JobNumberText = "Reference / Job #"
JobNumberTextWidth = Printer.TextWidth(JobNumberText)
Printer.CurrentX = (210-JobNumberTextWidth) / 1.5
Printer.CurrentY = VerticalMargin + 33
Printer.Print JobNumberText; Space(7);
Printer.CurrentY = VerticalMargin + 35
Printer.FontBold = False
Printer.FontSize = 10
Printer.Print jnum '###
Printer.FontName = "Arial"
Printer.FontSize = 12
Printer.FontBold = True
Printer.FontItalic = False
Printer.FontUnderline = False
Printer.FontStrikethru = False
Printer.ForeColor = RGB(0, 0, 0)
Printer.CurrentX = HorizontalMargin / 1.5
Col(0) = 10
Col(1) = 58
Col(2) = 106 'col width of 48mm(adjust to suit)
Col(3) = 154
NR = 53
For f = LBound(anystring) To UBound(anystring)
   '### anystring can be numbers, text,
   ' list box contents
 Printer.CurrentX = HorizontalMargin + (Col(k))
   'EG: 10 mm this time
   '58 mm next time etc.
Printer.CurrentY = VerticalMargin + (NR)
   'EG: 53 mm first time 60 mm next
    'then 67 mm ect.
Printer.Print anystring(f)
k = k + 1	'Next column on the next loop
If k = 4 Then NR = NR + 7: k = 0 'If you have 4
        'anystrings in
     ' this row, start a new row
If NR > 270 Then Printer.NewPage: NR = 20
      'Enough on this page
Next f			'Loop
Printer.EndDoc
ErrHandler:
	'User pressed Cancel button.
	EXIT SUB
End Sub
```

