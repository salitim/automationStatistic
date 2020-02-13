Option Explicit
On Error Resume Next
ExempleMacroExcel

Sub ExempleMacroExcel() 

  Dim ApplicationExcel 
  Dim ClasseurExcel 

  Set ApplicationExcel = CreateObject("Excel.Application") 
  Set ClasseurExcel = ApplicationExcel.Workbooks.Open("P:\test.xlsm") 
  
  ApplicationExcel.Visible = True   'les actions seront visibles. Pour tout lancer en arrière-plan, remplacer True par False
  ApplicationExcel.Run "Macro1" 'va lancer la macro "MacroTest1"
  ApplicationExcel.DisplayAlerts = False
  ApplicationExcel.Sheets("Feuil2").Delete
  ApplicationExcel.DisplayAlerts = True
  Set ClasseurExcel = Nothing 
  Set ApplicationExcel = Nothing 

End Sub