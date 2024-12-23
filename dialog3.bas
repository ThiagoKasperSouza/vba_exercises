Sub Main
	BasicLibraries.LoadLibrary("LibreMacro")
	DialogExample
End Sub

Sub	DialogExample
	Select case Dialog3("tem certeza?")
		case "Yes"
			Msgbox "clicou em sim"
		case "No"
			Msgbox "clicou em nao"
		case "Cancel"
			Msgbox "clicou em cancelar"
	End select
end sub
