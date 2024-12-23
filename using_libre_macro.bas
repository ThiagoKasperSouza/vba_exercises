Sub Main() 
  rem importacao libremacro
  BasicLibraries.LoadLibrary("LibreMacro")
  
  'escrever "teste na a3 de uma folha de calculo"
  Cell("Planilha1","A1").String = "teste"
  Cell("Planilha1","A2").Value = 2
  Cell("Planilha1","A3").Value = 5
  Cell("Planilha1","A4").Formula = "=A2+A3"
  Cell("Planilha1", REF(5,1)).String = "Oi"
  Dialogs
End Sub


Sub Dialogs
  'ConfirmDialog(pQuestion as String, Optional pDialogTitle as String) as Boolean
  ' QuestionDialog(pQuestion as String, Optional pDialogTitle as String)
  'RetryDialog(pQuestion as String, Optional pDialogTitle as String) as Boolean
  if QuestionDialog("tem certeza?") = true then
    msgbox("Operacao confirmada!")
  else
    msgbox("Operacao cancelada")
  end if
End Sub
