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
     Set dict = New Dicts

```
**_2. Load the dict_**
```

Call dict.load("src", 1, 2, 2, , dict.reg("^[a-zA-Z]{4}\s"), False, 100)

```
**_2. Loop through dict to see the result_**
```
    Dim k
    
    For Each k In dict.dict.keys
        Debug.Print k & "  " & dict.dict(k)
    Next k
```

## API
Placeholder

## Support or Contact
Having trouble with Pages? You can contact sgfxqw('__at__')gmail('__dot__')com and I will help you sort it out.
