Option Explicit
On Error Resume Next
ExempleMacroExcel

Sub ExempleMacroExcel() 

  Dim ApplicationExcel 
  Dim ClasseurExcel 

  Set ApplicationExcel = CreateObject("Excel.Application") 
  Set ClasseurExcel = ApplicationExcel.Workbooks.Open(ApplicationExcel.GetOpenFilename) 
  
  ApplicationExcel.Visible = True   'les actions seront visibles. Pour tout lancer en arrière-plan, remplacer True par False
  ApplicationExcel.Run "Decaler" 'va lancer la macro "Decaler"
  ApplicationExcel.DisplayAlerts = False
  ApplicationExcel.Sheets("statSVP").Delete
  ApplicationExcel.DisplayAlerts = True
  Set ClasseurExcel = Nothing 
  Set ApplicationExcel = Nothing 

End Sub 