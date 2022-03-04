Attribute VB_Name = "Modul1"
Function CAGR(Yt As Variant, Y0 As Variant, Period As Variant)
'Keyword Compound Annual Growth Rate (CAGR)
 CAGR = (Yt / Y0) ^ (1 / Period) - 1
End Function

