Sub Main
	BasicLibraries.LoadLibrary("LibreMacro")
	loopsExample
End Sub

Sub loopsExample
	for i=2 to 6 step 1
		if Cell("Planilha1", REF(i,2)).Value >=7 then
			 Cell("Planilha1", REF(i,3)).String = "Aprovado"
		else 
			Cell("Planilha1", REF(i,3)).String = "Reprovado"
		end if
	next
end sub
