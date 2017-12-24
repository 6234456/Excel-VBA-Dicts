## Introduction

**_Leverage the Power of Functional Programming in VBA_**
```

     Dim dict As New Dicts
     dict.map("_*2").filterKey("_>4")

```


## Basic Usage
_**Example :**_ 
* print out all the projects with a NPV > 5 to screen.

The spreadsheet "src" is as below.

```
       ![example1](http://qiou.eu/xl/example1.PNG "example1")
```

**_1. Create a new instance_**
```
     Dim d As New Dicts
```
**_2. Load the dict with data in the spreadsheet_**
```
     Call d.loadRng("", 1, d.rng(2, d.x("", 1)), 2)
```
**_3. filter it functionally and elegantly_**
```
     d.reduceRngX("{v}+{*}/1.1^({i}+1)").filterVal("{*}>5").p
```

## Support or Contact
Having trouble with this project? You can contact yang('__at__')qiou('__dot__')eu and I will help you sort it out.
