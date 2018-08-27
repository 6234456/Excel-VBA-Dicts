## Introduction

**_Leverage the Power of Functional Programming in VBA_**
```

     Dim dict As New Dicts
     dict.map("_*2").filter("_>4")

```


## Basic Usage
_**Example :**_ 
* print out all the projects with a NPV > 5 to screen (discount rate 10%).

The spreadsheet "src" is as below.

![example1](http://qiou.eu/xl/example1.PNG "example1")


**_1. Create a new instance_**
```
     Dim d As New Dicts
```
**_2. Load the dict with data in the spreadsheet_**
```
     With d.load("", 1, d.rng(2, d.x("", 1)), 2)
```
**_3. filter it functionally and elegantly_**
```
          .ranged("?+_/1.1^({i}+1)", AggregateMethod.AggReduce).filter("_>5").p
          
     End With
```

## Support or Contact
Having trouble with this project? You can contact yang('__at__')qiou('__dot__')eu and I will help you sort it out.
