## Introduction

**_Functional Programming in VBA_**
```

     Dim dict As New Dicts
     dict.map("_*2").filterKey("_>4")

```


## Basic Usage
_**Example :**_ 
* to load all the invoices with invoice number starting with 4 alphabetical letters on the spreadsheet "src"
* to set the default pieces to be 100

The spreadsheet "src" is as below.

```
             A                      B

     1    Invoice Number          Pieces

     2    RELD 12323              1400

     3    RE 12324                500

     4    RELD 12325          

     5    RELD 12326              100
```

**_1. Create a new instance_**
```

     Dim dict As New Dicts

```
**_2. Load the dict_**
```

Call dict.load("src", 1, 2, 2, , dict.reg("^[a-zA-Z]{4}\s"), False, 100)

```
**_3. Loop through dict to print the result_**
```
    dict.p
```

## API


###**load**###
_to load the range into dict_

Parameters:
```
Byval targSht As String                             'name of the target Sheet. Empty string if target sheet is current sheet.
ByVal targKeyCol As Integer                         'the column number of the keys in the dictionary
ByVal targValCol                                    'the column number of the values in the dictionary
Optional targRowBegine As Variant                   'the first row of the range, default to be 1
Optional ByVal targRowEnd As Variant                'the last row of the range, default to be the row where key column ends
Optional ByVal reg As Variant                       'the regular expression to filter the keys
Optional ByVal ignoreNullVal As Boolean             'ignore if the value is null
Optional ByVal setNullValto As Variant              'if ignoreNullVal is false, set null to this value
```

## Support or Contact
Having trouble with this project? You can contact sgfxqw('__at__')gmail('__dot__')com and I will help you sort it out.
