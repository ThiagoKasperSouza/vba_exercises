Sub Main() 
   IntOperations
   ArithmeticOps
End Sub

Sub IntOperations()
	dim resultado as Integer
	' resultado soma
	resultado = 8+2
	msgbox "resultado soma: " & resultado 
	'resultado subtracao
	resultado = 8-2
	msgbox "resultado subtracao: " & resultado
	resultado = 8*2
	msgbox "resultado multiplicacao: " & resultado
	resultado = 8/2 
	msgbox "resultado divisao: " & resultado 
End Sub

Sub ArithmeticOps()
    Dim expression As String
    Dim operand1 As Double
    Dim operand2 As Double
    Dim operator As String
    Dim result As Double
    Dim pos As Integer

    ' Solicita ao usuário que insira uma expressão
    expression = InputBox("Digite uma operaçao entre 2 algarismos (ex: 10 + 5):")

    ' Encontra a posição do operador
    pos = InStr(expression, "+")
    If pos = 0 Then pos = InStr(expression, "-")
    If pos = 0 Then pos = InStr(expression, "*")
    If pos = 0 Then pos = InStr(expression, "/")

    ' Verifica se um operador foi encontrado 
    If pos = 0 Then
        MsgBox "Operador inválido! Por favor, use +, -, * ou /."
        Exit Sub
    End If

    ' Separa os operandos e o operador
    operator = Mid(expression, pos, 1)
    operand1 = Val(Trim(Left(expression, pos - 1)))
    operand2 = Val(Trim(Mid(expression, pos + 1)))

    ' Realiza a operação com base no operador
    Select Case operator
        Case "+"
            result = operand1 + operand2
        Case "-"
            result = operand1 - operand2
        Case "*"
            result = operand1 * operand2
        Case "/"
            If operand2 <> 0 Then
                result = operand1 / operand2
            Else
                MsgBox "Erro: Divisão por zero!"
                Exit Sub
            End If
    End Select

    ' Exibe o resultado
    MsgBox "O resultado de " & expression & " é: " & result
End Sub
