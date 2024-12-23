Sub Main() 
	rem importacao libremacro
	BasicLibraries.LoadLibrary("LibreMacro")
	Conditions
	

End Sub


Sub Conditions
	teste = InputBox("Digite o nome do aluno: ")
	
	Select case teste 
		Case "Ana"
			checkGrade("B2", "C2")
		Case "Julio"
			checkGrade("B3", "C3")
		Case "Pedro"
			checkGrade("B4", "C4")
		Case "Marcos"
			checkGrade("B5", "C5")
		Case "Paula"
			checkGrade("B6", "C6")
	End select
			
End Sub

Sub checkGrade(p1 as String, p2 as String)
	if Cell("Planilha1", p1).Value >= 7 then
		Cell("Planilha1", p2).String = "Aprovado"
		Msgbox "Aprovado"
	else
		Cell("Planilha1", p2).String = "Reprovado"
		Msgbox "Reprovado"
	end if
End sub
