Attribute VB_Name = "vgTexto"
Option Explicit

'******************************************************************************
'FIND TEXT BETWEEN TWO STRINGS
'This function is useful for returning a portion of a
'string between two points in the string.
'You could, for example, extract a range name returned by
'Excel found between parentheses.
'Note that it only works for the first occurrences
'of the start and stop delimiters.
'******************************************************************************
Public Function Between(sText As String, _
                        sStart As String, _
                        sEnd As String) As String
Dim lLeft As Long, lRight As Long

    lLeft = InStr(sText, sStart) + (Len(sStart) - 1)
    lRight = InStr(lLeft + 1, sText, sEnd)

    If lRight > lLeft Then Between = _
        Mid$(sText, lLeft + 1, ((lRight - 1) - lLeft))
End Function

Public Function Acerta_Plic(fTexto As String) As String
  
    If InStr(fTexto, Chr(39)) Then
        Acerta_Plic = Replace(fTexto, Chr(39), Chr(39) & Chr(39))
    Else
        Acerta_Plic = fTexto
    End If

End Function

'******************************************************************************
'* Fun��o : EXISTE_PLIC                                                       *
'* Descri��o : verifica a existencia de plics em um texto                     *
'* Parametros :                                                               *
'*      fTexto - descri��o                                                    *
'* Sa�da :                                                                    *
'* true - ok, n�o h� plics                                                    *
'* false - erro, h� plics no texto                                            *
'******************************************************************************
Public Function EXISTE_PLIC(fTexto As String) As Boolean
 
    EXISTE_PLIC = True
    
    If InStr(fTexto, "'") > 0 Then
        EXISTE_PLIC = False
        MsgBox "N�o � permitido utilizar o caracter ' no texto.", vbInformation
    End If

End Function



'******************************************************************************
'* Fun��o :  ImprimeNota                                                      *
'* Descri��o : formata texto em blocos de 79 caracteres                       *
'* Parametros :                                                               *
'*      Nota - texto a ser formatado                                          *
'* Sa�da :                                                                    *
'*      texto formatado                                                       *
'******************************************************************************
Public Function ImprimeNota(ByVal Nota As String) As String
Dim Ini, Tam As Byte
Dim Campo As String

    Ini = 1
    Tam = 79

    Tam = InStr(Ini, Trim$(Mid$(Nota, Ini, Tam)), Chr(13))
    If Tam = 0 Then Tam = 79

    Campo = Trim$(Mid$(Nota, Ini, Tam))
    
    Ini = Ini + Tam
    While Len(Trim$(Mid$(Nota, Ini, Tam))) > 0
       Tam = InStr(Ini, Trim$(Mid$(Nota, Ini, Tam)), Chr(13))
       If Tam = 0 Then Tam = 79
       
       Campo = Campo + Chr(13) + Trim$(Mid$(Nota, Ini, Tam))
       Ini = Ini + Tam
    Wend

    ImprimeNota = Campo

End Function

'******************************************************************************
'recebe string alfanum�rica
'retorna string de letras mai�sculas  sem pontos, v�rgulas, n�meros etc.
'exemplo FU_LimpaNumero("Adq-7465") = "ADQ"
'******************************************************************************
Public Function LimpaAlfa(Campo As String) As String
Dim VA_Posicao As Integer
Dim VA_Caracter As String * 1
Dim VA_Resultado As String
    
    VA_Resultado = ""
    VA_Posicao = 1
    Campo = UCase(Campo)
    
    Do While VA_Posicao <= Len(Campo)
        VA_Caracter = Mid$(Campo, VA_Posicao, 1)
        If Asc(VA_Caracter) > 64 And Asc(VA_Caracter) < 91 Then
            VA_Resultado = VA_Resultado & VA_Caracter
        End If
        VA_Posicao = VA_Posicao + 1
    Loop
    
    LimpaAlfa = VA_Resultado

End Function

'******************************************************************************
'recebe string num�rica
'retorna string num�rica sem pontos, v�rgulas etc.
'exemplo FU_LimpaNumero("1.245,90") = "1234590"
'******************************************************************************
Public Function LimpaNumero(Campo As String) As String
Dim VA_Posicao As Integer
Dim VA_Caracter As String * 1
Dim VA_Resultado As String
    
    VA_Resultado = ""
    VA_Posicao = 1

    Do While VA_Posicao <= Len(Campo)
        VA_Caracter = Mid$(Campo, VA_Posicao, 1)
        If IsNumeric(VA_Caracter) Then
            VA_Resultado = VA_Resultado & VA_Caracter
        End If
        VA_Posicao = VA_Posicao + 1
    Loop

    LimpaNumero = VA_Resultado

End Function

'******************************************************************************
'* Fun��o : Limpa_Plic                                                        *
'* Descri��o : retira plics de um texto                                       *
'* Parametros :                                                               *
'*      fTexto - texto a ser verificado                                       *
'* Sa�da :                                                                    *
'*      texto sem plics                                                       *
'******************************************************************************
Public Function Limpa_Plic(fTexto As String) As String
Dim i As Integer
    
    Limpa_Plic = ""
    For i = 1 To Len(fTexto)
        If Mid(fTexto, i, 1) <> "'" Then
            Limpa_Plic = Limpa_Plic & Mid(fTexto, i, 1)
        End If
    Next i

End Function

'******************************************************************************
'* Fun��o : LPSTRtoVBString$                                                  *
'* Descri��o : extrai um string vb de um buffer contendo um string terminado  *
'* por nulo (LPSTR)                                                           *
'* Parametros :                                                               *
'*      s$ - string a converter (LPSTR)                                       *             '* Sa�da :                                                                    *
'*      string vb                                                             *
'******************************************************************************
Public Function LPSTRtoVBString$(ByVal s$)
' Extracts a VB string from a buffer containing a null terminated string
Dim nullpos&

    nullpos& = InStr(s$, Chr$(0))
    If nullpos > 0 Then
        LPSTRtoVBString = Left$(s$, nullpos - 1)
    Else
        LPSTRtoVBString = ""
    End If

End Function

'******************************************************************************
'* Fun��o : NAO_TECLA_PLIC                                                    *
'* Descri��o : filtra o uso da tecla <'>                                      *
'* Parametros :                                                               *
'*      fKeyAscii - c�digo ascii do caracter a filtrar                        *
'* Sa�da :                                                                    *
'*      c�digo ascii de um caracter diferente de plic <'>                     *
'* Cuidado: N�o permite tamb�m acento agudo                                   *
'******************************************************************************
Public Function NAO_TECLA_PLIC(fKeyAscii As Integer) As Integer
    
    If fKeyAscii = 39 Then
        NAO_TECLA_PLIC = vbNull
    Else
        NAO_TECLA_PLIC = fKeyAscii
    End If

End Function

'******************************************************************************
'A BETTER USE FOR STRCONV
'properName = StrConv(text, vbProperCase)
'Be aware that this variant of StrConv also forces a
'conversion to lowercase for all the characters not at
'the beginning of a word.
'In other words, "seattle, USA," is converted to
'"Seattle, Usa," which this function doesn't do.
'******************************************************************************
Public Function ProperCase(text As String) As String
Dim result As String, i As Integer
    
    result = StrConv(text, vbProperCase)
    ' restore all those characters that
    ' were uppercase in the original string
    For i = 1 To Len(text)
        Select Case Asc(Mid$(text, i, 1))
        Case 65 To 90       ' A-Z
            Mid$(result, i, 1) = Mid$(text, i, 1)
        End Select
    Next
    ProperCase = result

End Function

'******************************************************************************
'GENERATE RANDOM STRINGS
'This code helps test SQL functions or other
'string-manipulation routines so you can generate random
'strings. You can generate random-length strings with
'random characters and set ASCII bounds, both upper and
'lower:
'******************************************************************************
Public Function RandomString(iLowerBoundAscii As Integer, _
                             iUpperBoundAscii As Integer, _
                             lLowerBoundLength As Long, _
                             lUpperBoundLength As Long) As String
Dim sHoldString As String
Dim lLength As Long
Dim lCount As Long

    'Verify boundaries
    If iLowerBoundAscii < 0 Then iLowerBoundAscii = 0
    If iLowerBoundAscii > 255 Then iLowerBoundAscii = 255
    If iUpperBoundAscii < 0 Then iUpperBoundAscii = 0
    If iUpperBoundAscii > 255 Then iUpperBoundAscii = 255
    If lLowerBoundLength < 0 Then lLowerBoundLength = 0

    'Set a random length
    lLength = Int((CDbl(lUpperBoundLength) - _
        CDbl(lLowerBoundLength) + _
        1) * Rnd + lLowerBoundLength)

    'Create the random string
    For lCount = 1 To lLength
        sHoldString = sHoldString & _
            Chr(Int((iUpperBoundAscii - iLowerBoundAscii _
            + 1) * Rnd + iLowerBoundAscii))
    Next
    RandomString = sHoldString

End Function

Public Function StringCount(sText As String, _
                            sFind As String) As Long
Dim lFind As Long
Dim lLast As Long

    Do
        lFind = InStr(lLast + 1, sText, sFind)
        If lFind Then
            lLast = lFind
            StringCount = StringCount + 1
        End If
    Loop Until lFind = 0

End Function

'******************************************************************************
'* Fun��o : TECLA_LETRA_NUMERO                                                *
'* Descri��o : filtra caracteres n�o alfanumericos (simbolos e caracteres     *
'* especiais)                                                                 *
'* Parametros :                                                               *
'*      fKeyAscii - c�digo do caracter a filtrar                              *
'*      fNPermiteEspaco - flag que autoriza a utiliza��o do caracter <espa�o> *
'* Sa�da :                                                                    *
'*      c�digo ascii contendo apenas caracteres alfanumericos                 *
'******************************************************************************
Public Function TECLA_LETRA_NUMERO(fKeyAscii As Integer, _
                                   Optional fNPermiteEspaco As Boolean) As Integer
    
    '48 A 57 s�o 0 - 9
    '65 A 90 s�o A - Z
    '97 A 122 s�o a - z
    
    If fNPermiteEspaco And fKeyAscii = vbKeySpace Then
        TECLA_LETRA_NUMERO = vbNull
        Exit Function
    End If
    
    If (fKeyAscii >= 48 And fKeyAscii <= 57) Or _
       (fKeyAscii >= 65 And fKeyAscii <= 90) Or _
       (fKeyAscii >= 97 And fKeyAscii <= 122) Or _
        fKeyAscii = vbKeySpace Or _
        fKeyAscii = vbKeyBack Or _
        fKeyAscii = vbKeyDelete Then
    '3 A 26 s�o ctrl C, X, V, Z
'        fKeyAscii = 3 Or _
'        fKeyAscii = 22 Or _
'        fKeyAscii = 24 Or _
'        fKeyAscii = 26 Or _
'        fKeyAscii = vbKeyControl Or _

        TECLA_LETRA_NUMERO = fKeyAscii
    Else
        TECLA_LETRA_NUMERO = vbNull
    End If

End Function

'******************************************************************************
'* Fun��o : TESTA_PLIC                                                        *
'* Descri��o : verifica a existencia de plics <'> em um controle              *
'* Parametros :                                                               *
'*      fCampo - controle a verificar                                         *
'* Sa�da :                                                                    *
'*      true se o controle possui plics                                       *
'*      false se o controle n�o possui plics                                  *
'******************************************************************************
Public Function TESTA_PLIC(fCampo As Control) As Boolean

    TESTA_PLIC = False
    
    If InStr(fCampo.text, "'") > 0 Then
        TESTA_PLIC = True
        MsgBox "N�o � permitido usar o caracter ' neste campo.", vbInformation
        fCampo.SetFocus
    End If
    
End Function

Public Function TIRA_ACENTO_LETRA(fLetra As String) As String
Dim NovaLetra As String

    Select Case fLetra
        Case "�"
            NovaLetra = "a"
        Case "�"
            NovaLetra = "e"
        Case "�"
            NovaLetra = "i"
        Case "�"
            NovaLetra = "o"
        Case "�"
            NovaLetra = "u"

        Case "�"
            NovaLetra = "a"
        Case "�"
            NovaLetra = "e"
        Case "�"
            NovaLetra = "i"
        Case "�"
            NovaLetra = "o"
        Case "�"
            NovaLetra = "u"

        Case "�"
            NovaLetra = "a"
        Case "�"
            NovaLetra = "e"
        Case "�"
            NovaLetra = "i"
        Case "�"
            NovaLetra = "o"
        Case "�"
            NovaLetra = "u"

        Case "�"
            NovaLetra = "a"
        Case "�"
            NovaLetra = "e"
        Case "�"
            NovaLetra = "i"
        Case "�"
            NovaLetra = "o"
        Case "�"
            NovaLetra = "u"

        Case "�"
            NovaLetra = "a"
        Case "�"
            NovaLetra = "o"

        Case "�"
            NovaLetra = "c"

        Case "�"
            NovaLetra = "A"
        Case "�"
            NovaLetra = "E"
        Case "�"
            NovaLetra = "I"
        Case "�"
            NovaLetra = "O"
        Case "�"
            NovaLetra = "U"
        
        Case "�"
            NovaLetra = "A"
        Case "�"
            NovaLetra = "E"
        Case "�"
            NovaLetra = "I"
        Case "�"
            NovaLetra = "O"
        Case "�"
            NovaLetra = "U"
        
        Case "�"
            NovaLetra = "A"
        Case "�"
            NovaLetra = "E"
        Case "�"
            NovaLetra = "I"
        Case "�"
            NovaLetra = "O"
        Case "�"
            NovaLetra = "U"
        
        Case "�"
            NovaLetra = "A"
        Case "�"
            NovaLetra = "E"
        Case "�"
            NovaLetra = "I"
        Case "�"
            NovaLetra = "O"
        Case "�"
            NovaLetra = "U"
        
        Case "�"
            NovaLetra = "A"
        Case "�"
            NovaLetra = "O"
        
        Case "�"
            NovaLetra = "C"
            
        Case Else
            NovaLetra = fLetra
    End Select
    
    TIRA_ACENTO_LETRA = NovaLetra
    
End Function

'*******************************************************************************
'* Fun��o : TIRA_ZEROS_ESQUERDA                                                *            '* Descri��o : retira os zeros a esquerda de um string e emite uma mensagem de *
'* se desejado                                                                 *
'* Parametros :                                                                *
'*      Texto - valor a ser alterado                                           *
'*      Msg - mensagem de erro caso existam zeros a esquerda                   *
'* Sa�da :                                                                     *
'*      string sem zeros a esquerda                                            *
'*******************************************************************************
Public Function TiraZerosAEsquerda(Texto As String, _
                                   Optional msg As Boolean = True) As String
Dim i As Integer
    
    TiraZerosAEsquerda = Trim$(Texto)
    
    For i% = 1 To Len(TiraZerosAEsquerda)
        If Left$(TiraZerosAEsquerda, 1) = "0" Then
            TiraZerosAEsquerda = Mid$(TiraZerosAEsquerda, 2, Len(TiraZerosAEsquerda))
        Else
            If msg And TiraZerosAEsquerda <> Trim$(Texto) Then
                MsgBox "N�o � permitido informar ZEROS a esquerda.", vbInformation
            End If
            Exit Function
        End If
    Next i%
    
End Function

'******************************************************************************
'* Fun��o : TROCA_PLIC_ASPAS                                                  *
'* Descri��o : usar no evento Keypress                                        *
'******************************************************************************
Public Function TROCA_PLIC_ASPAS(fKeyAscii As Integer) As Integer
    
    If fKeyAscii = 39 Then
        TROCA_PLIC_ASPAS = 34
    Else
        TROCA_PLIC_ASPAS = fKeyAscii
    End If

End Function

'******************************************************************************
'* Fun��o : Upper                                                             *
'* Descri��o : converte um c�digo ascii de um caracter para o c�digo          *
'* correspondente em ma�sculas                                                *
'* Parametros :                                                               *
'*      fKeyAscii - c�digo ascii referente a um caracter                      *
'* Sa�da :                                                                    *
'*      c�digo ascii do caracter em mai�sculas                                *
'******************************************************************************
Public Function Upper(fKeyAscii As Integer) As Integer

    fKeyAscii = Asc(UCase$(Chr$(fKeyAscii)))
    
    Upper = fKeyAscii
    
End Function

Public Function FormataTexto(fTexto As String, _
                             Optional fCaption As Boolean = True) As String
    
    FormataTexto = TrocaPlicAspas(Trim$(fTexto))

    If fCaption Then
        FormataTexto = UCase$(FormataTexto)
    End If
    
End Function

Public Function TrocaPlicAspas(fTexto As String) As String
    
    TrocaPlicAspas = Replace(fTexto, Chr(39), Chr(34))

End Function




