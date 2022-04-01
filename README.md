# User defined Functions and Macros VBA
Funtions and Macros that can be used in excel to improve the usability of the software


- [SumByFontColor()](https://github.com/carloscastillom/User_defined_Functions_VBA/blob/main/SumbyFontColor.bas)

  Function that sums the value of a cell of an specific color Font. In case you need a different condition please modify it. 

- [RefresAllPivot()](https://github.com/carloscastillom/User_defined_Functions_VBA/blob/main/RefreshAllPivotTables.bas)

  Macro that updates all the Pivot tables in the workbook.

- [CAGR(Yt: Final Value, Y0: Initial Value, T: Period)](https://github.com/carloscastillom/User-defined-Functions-Macros-VBA/blob/main/CAGR.bas) 

  Function that calculates the compound anual growth rate between two values and the amount of periods (Yt: Final Value, Y0: Initial Value, T: Period)
  
  ```
  Function CAGR(Yt As Variant, Y0 As Variant, T As Variant)
  'Keyword Compound Annual Growth Rate (CAGR)
  CAGR = (Yt / Y0) ^ (1 / T) - 1
  End Function
  ```

