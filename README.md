<div align="center">

## Round\(\) Vs Format\(\)


</div>

### Description

If your application involve ROUNDING NUMBER than you MUST READ this!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ken Wong](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ken-wong.md)
**Level**          |Beginner
**User Rating**    |4.3 (34 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ken-wong-round-vs-format__1-35726/archive/master.zip)





### Source Code

```
Rounding number is important fuction especially if you doing some Accounting, Math, etc calculation. Check the following fact.
Let say, X = 2.344 + 2.341, there for, X = 4.685
Round X to 2 decimal point = 4.68
Refer to IEEE standard,
4.685 rounded to 2 decimal point is 4.68
4.675 rounded to 2 decimal point is 4.68
The problem when you use Round():
Fact 1 : Round(2.344 + 2.341, 2) = 4.69
Fact 2: Round(4.685, 2) = 4.68
You can see the different answer between Fact 1 & 2. In accounting point of view, the 0.01 cent must accountable.
Recommandation:
Use Format() instead of Round()
Example:
Format(2.344 + 2.341, "#.00") = 4.68
Format(4.685, "#.00") = 4.68
Hope this will help.
```

