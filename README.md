# Data

## Extract URL From link using Microsoft Excel
- Using a VBA (Visual Basic for Applications) custom function:
```
    Press Alt + F11 to open the VBA editor.
    Click "Insert" > "Module".
```

- Copy and paste this VBA code into the module window: 
    vba
```
    Function GetURL(Hlink As Range) As String
        On Error Resume Next
        GetURL = Hlink.Hyperlinks(1).Address
    End Function
```
- Close the VBA editor.
- Now, in any cell on your worksheet, 
you can use the formula =GetURL(A1) (replacing "A1" with the cell containing the hyperlink) to extract the URL. 

