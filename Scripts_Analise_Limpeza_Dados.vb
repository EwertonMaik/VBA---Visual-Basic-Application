' Scripts VBA - Visual Basic Application -
' Exempos utilizado para analise e limpeza de dados

'Funcao para converter os caracteres string para UTF8
'Funcao Publica que recebe parametro de uma variavel string
Public Function fnConverterUTF8(ByVal Texto_para_converter As String)
	
	'Criacao e Definicao das Variaveis utilizadas na Funcao
    Dim l As Long, sUTF8 As String
    Dim iChar As Integer
    Dim iChar2 As Integer
    
	'Metodo que torna a funcao volatil, recalcula automaticamente sempre que houver edicao na celula
    Application.Volatile
    
	'Estrutura de repeticao FOR para ler e percorrer toda extensao da String, utiliza LEN para obter tamanho da String
    For l = 1 To Len(Texto_para_converter)
		
		'Funcao MID retorna um pedaco da String, neste caso, apenas um caracter, baseado na posicao da variavel FOR percorrida
		'Funcao ASC retorna o codigo do caractere em ASC do 1º caractere de uma String, neste caso, o obtido pela funcao MID
        iChar = Asc( Mid(Texto_para_converter, l, 1) )
		
		'Estrutura condicional IF para validar se o tamanho da String e maior que 127
        If iChar > 127 Then
            If Not iChar And 32 Then
            iChar2 = Asc( Mid(Texto_para_converter, l + 1, 1))
            sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
            l = l + 1
        Else
            Dim iChar3 As Integer
            iChar2 = Asc(Mid(Texto_para_converter, l + 1, 1))
            iChar3 = Asc(Mid(Texto_para_converter, l + 2, 1))
            sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
            l = l + 2
        End If
            Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    Next l
	
    fnConverterUTF8 = sUTF8
	
End Function


' Funcao para remover tipos de acentuacao de uma string
Function fnRetirarAcentos(ByVal vStrPalavra As String) As String
    Dim lstrEspecial    As String
    Dim lstrSubstituto  As String
    Dim lstrAlterada    As String
    Dim liControle      As Integer
    Dim liPosicao       As Integer
    Dim lstrLetra       As String
    
    Application.Volatile

    lstrEspecial = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
    lstrSubstituto = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
 
    lstrAlterada = ""
 
    If vStrPalavra <> "" Then
        For liControle = 1 To Len(vStrPalavra)
            lstrLetra = Mid(vStrPalavra, liControle, 1)
            liPosicao = InStr(lstrEspecial, lstrLetra)
        
            If liPosicao > 0 Then
                lstrLetra = Mid(lstrSubstituto, liPosicao, 1)
            End If
        
            lstrAlterada = lstrAlterada & lstrLetra
        Next
        
        fnRetirarAcentos = lstrAlterada
    End If
End Function