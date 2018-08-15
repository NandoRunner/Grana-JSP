Attribute VB_Name = "vgNumero"
Option Explicit

'----------------------------------------
' Declaração de APIs e Constantes Locais
'----------------------------------------
'FormatoNumeroValido()
Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SDECIMAL = &HE         'decimal separator
Private Const LOCALE_SMONTHOUSANDSEP = &H17 'monetary thousand separator
Private Const LOCALE_SMONDECIMALSEP = &H16  'monetary decimal separator
Private Const LOCALE_STHOUSAND = &HF        'thousand separator
Private Const LOCALE_SCURRENCY = &H14       'currency symbol
Private Const LOCALE_SNEGATIVESIGN = &H51   'negative sign

Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
       (ByVal Locale As Long, _
        ByVal LCType As Long, _
        ByVal lpLCData As String, _
        ByVal cchData As Long) As Long

'----------------------------------------
' Declaração de Constantes Globais
'----------------------------------------
'Formatos do Painel de Controle
Public Enum gTpNumFormat
    US_NumFormat = 0
    BR_NumFormat = 1
End Enum

Public Function FormataMoeda(fValor, _
                             Optional fTestaFormato As Boolean = False, _
                             Optional fMaxVal As Currency = 922337203685477#, _
                             Optional fSimboloMoeda As String = "")
    
    'Formata de acordo com as configurações do Painel de Controle do Windows
    'Usar fTestaFormato caso não se possa garantir o formato de fValor (ex.: entrada de dados)
    
    'OBS.: Uma moeda de exibição diferente da configurada no Painel de Controle do Windows
    '      pode, alternativamente, ser informada pelo usuário

    If fValor = "" Then
        FormataMoeda = fValor
        Exit Function
    End If
    
    If fSimboloMoeda = "" Then
        fSimboloMoeda = VerificaSimboloMoeda()
    End If
    
    fSimboloMoeda = Trim$(fSimboloMoeda)
    
    fValor = FormataNumeroDecimal(fValor, True, fTestaFormato, 2)
    
    MaxCurrency fValor, True, 2, fMaxVal
    
    fValor = Replace(fValor, fSimboloMoeda, "")
    
    If Trim$(fValor) = "-" Then
        fValor = "0.00"
    End If
        
    FormataMoeda = fSimboloMoeda & " " & fValor
    
End Function

Public Function FormataNumero(fValor, _
                              Optional fTestaFormato As Boolean = False)
Dim Valor

    'Formata de acordo com as configurações do Painel de Controle do Windows

    If Not FormatoNumeroValido() Then
        'vgGlobal.Fim
        
        Exit Function
    End If
    
    If fValor = "" Then
        FormataNumero = fValor
        Exit Function
    End If
    
    FormataNumero = Format(fValor, "#,##0")

End Function

Public Function FormataNumeroDecimal(fValor, _
                                     Optional fUsaAgrupDigitos As Boolean = True, _
                                     Optional fTestaFormato As Boolean = False, _
                                     Optional fPrecisao As Byte = 2)
Dim Formato As String
Dim Valor

    'Formata de acordo com as configurações do Painel de Controle do Windows
    'Usar fTestaFormato caso não se possa garantir o formato de fValor (ex.: entrada de dados)

    If Not FormatoNumeroValido() Then
        'vgGlobal.Fim
        Exit Function
    End If
    
    Valor = Trim(Replace(fValor, "%", ""))
    
    If Valor = "" Then
        FormataNumeroDecimal = Valor
        Exit Function
    End If
    
    If fTestaFormato Then
        Valor = GaranteUnicoSepDecimal(Valor, fPrecisao)
        
        Valor = TrocaSeparador(Valor, VerificaFormatoNumericoBRUS())
    End If
    
    If fUsaAgrupDigitos Then
        Formato = "#,##0"
    Else
        Formato = "#0"
    End If
    
    If fPrecisao > 0 Then
        Formato = Formato & "." & String(fPrecisao, "0")
    End If
    
    FormataNumeroDecimal = Format(Valor, Formato)

End Function

Public Function FormataNumeroDecimalGravacao(fValor, _
                                             Optional fFormatoNumerico As gTpNumFormat = US_NumFormat, _
                                             Optional fTestaFormato As Boolean = True)
    
    '----------------------------------------------------------------------
    'ATENÇÃO: Não utilizar no caso de números inteiros (sem casas decimais)
    '----------------------------------------------------------------------
    'Troca o formato dos separadores de acordo com a escolha de fFormatoNumerico
    'Os separadores para fFormatoNumerico são obtidos do painel de controle do windows
    'Usar fTestaFormato caso não se possa garantir o formato de fValor (ex.: entrada de dados)
    
    If fValor = "" Then
        FormataNumeroDecimalGravacao = fValor
        Exit Function
    End If
    
    FormataNumeroDecimalGravacao = FormataNumeroDecimal(fValor, False, fTestaFormato)
    
    FormataNumeroDecimalGravacao = TrocaSeparador(FormataNumeroDecimalGravacao, fFormatoNumerico)
    
End Function

Public Function FormataPercent(fValor, _
                               Optional fUsaAgrupDigitos As Boolean = True, _
                               Optional fTestaFormato As Boolean = False, _
                               Optional fPrecisao As Byte = 2)
Dim Formato As String
Dim Valor

    'Formata de acordo com as configurações do Painel de Controle do Windows
    'Usar fTestaFormato caso não se possa garantir o formato de fValor (ex.: entrada de dados)
    
    If Not FormatoNumeroValido() Then
        'vgGlobal.Fim
        Exit Function
    End If
    
    Valor = Trim(Replace(fValor, "%", ""))
    
    If Valor = "" Then
        FormataPercent = Valor
        Exit Function
    End If
    
    If fTestaFormato Then
        Valor = GaranteUnicoSepDecimal(Valor, fPrecisao)
    
        Valor = TrocaSeparador(Valor, VerificaFormatoNumericoBRUS())
    End If
    
    Valor = Valor / 100
    
    If fUsaAgrupDigitos Then
        Formato = "#,##0"
    Else
        Formato = "#0"
    End If
    
    If fPrecisao > 0 Then
        Formato = Formato & "." & String(fPrecisao, "0")
    End If
    
    Formato = Formato & " %"
    
    FormataPercent = Format(Valor, Formato)

End Function

'******************************************************************************
'* Função : FormatoNumeroValido                                               *
'* Descrição :                                                                *
'*      Obtém o formato definido no Control Panel através da função           *
'*      GetLocaleInfo)                                                        *
'*      Se formato para número e moeda é diferente                            *
'*      então Formato é inválido                                              *
'* Saída :                                                                    *
'*      true - formato numerico compativel com o sistema                      *
'*      false - formato numerico incompatível com o sistema                   *
'******************************************************************************
Public Function FormatoNumeroValido(Optional fChamaPainelControle As Boolean = True) As Boolean
Dim FormData As String
Dim msg As String

    FormatoNumeroValido = True
    
    If VerificaSimboloNegativo() <> "-" Then
        FormatoNumeroValido = False
        
        msg = "Não é possível continuar." & vbCrLf & vbCrLf & _
              "Motivo: Símbolo de Sinal Negativo inválido." & vbCrLf & vbCrLf & _
              "Alterar o Símbolo de Sinal Negativo para '-'," & vbCrLf & _
              "na Configuração Regional de Números do Painel de Controle."
    Else
        FormData = VerificaFormatoNumero() & VerificaFormatoMoeda()
            
        If FormData <> ",.,." And FormData <> ".,.," Then
            FormatoNumeroValido = False
            
            msg = "Não é possível continuar." & vbCrLf & vbCrLf & _
                  "Motivo: Símbolos de Separadores Numéricos inválidos ou inconsistentes." & vbCrLf & vbCrLf & _
                  "Alterar a Configuração Regional de Símbolo Agrupador de Dígitos e de Símbolo Decimal," & vbCrLf & _
                  "para '.' e ',' ou para ',' e '.', igualmente para Número e Moeda no Painel de Controle."
        End If
    End If
    
    If Not FormatoNumeroValido Then
        If fChamaPainelControle Then
            MsgBox msg, vbCritical
            
            On Error Resume Next
            Shell ("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,," & 1)
            Err = 0
        End If
    End If
 
End Function

'******************************************************************************
'* Função : FormatoNumeroValidoBRUS                                           *
'* Descrição :                                                                *
'*      Obtém o formato definido no Control Panel                             *
'*          através da função GetLocaleInfo)                                  *
'*      Se formato, tanto para números quanto moeda, não é "." para separador *
'*      de milhares e "," para separador decimal então Formato é inválido     *
'* Saída :                                                                    *
'*      true - separadores de digitos compativel com o sistema  : '.'         *
'*      false - separadores de digitos incompatíveis com o sistema            *
'******************************************************************************
Public Function FormatoNumeroValidoBRUS(fFormatoNumerico As gTpNumFormat, _
                                        Optional fChamaPainelControle As Boolean = True) As Boolean
Dim FormData As String
Dim AgrupDigitos As String
Dim SepDecimal As String
Dim msg As String

    FormatoNumeroValidoBRUS = True
    
    If VerificaSimboloNegativo() <> "-" Then
        FormatoNumeroValidoBRUS = False
        
        msg = "Não é possível continuar." & vbCrLf & vbCrLf & _
              "Motivo: Símbolo de Sinal Negativo inválido." & vbCrLf & vbCrLf & _
              "Alterar o Símbolo de Sinal Negativo para '-'," & vbCrLf & _
              "na Configuração Regional de Números do Painel de Controle."
    Else
        FormData = VerificaFormatoNumero() & VerificaFormatoMoeda()
            
        If fFormatoNumerico = BR_NumFormat Then
            AgrupDigitos = "."
            SepDecimal = ","
        Else
            AgrupDigitos = ","
            SepDecimal = "."
        End If
        
        If FormData <> AgrupDigitos & SepDecimal & AgrupDigitos & SepDecimal Then
            FormatoNumeroValidoBRUS = False
            
            If fChamaPainelControle Then
                msg = "Não é possível continuar." & vbCrLf & vbCrLf & _
                      "Motivo: Símbolos de Separadores Numéricos inválidos ou inconsistentes." & vbCrLf & vbCrLf & _
                      "Alterar a Configuração Regional de Símbolo Agrupador de Dígitos para '" & AgrupDigitos & "'" & vbCrLf & _
                      "e de Símbolo Decimal para '" & SepDecimal & "', para Número e Moeda no Painel de Controle."
            End If
        End If
    End If
    
    If Not FormatoNumeroValidoBRUS Then
        If fChamaPainelControle Then
            MsgBox msg, vbCritical
            
            On Error Resume Next
            Shell ("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,," & 1)
            Err = 0
        End If
    End If
 
End Function

Public Function GaranteUnicoSepDecimal(fValor, _
                                       Optional fPrecisao As Byte = 2)
Dim Valor
Dim SepDecimal As String

    'Limpa se houver mais de um separador decimal
    
    On Error GoTo Erro_Overflow
    
    Valor = fValor
    
    SepDecimal = VerificaSepDecimalNumero()
    
    If vgTexto.StringCount(CStr(Valor), SepDecimal) > 1 Then
        If InStr(Valor, SepDecimal) > 0 Then
            Valor = Replace(Valor, SepDecimal, "")
            
            Valor = CCur(Valor) / Val("1" & String(fPrecisao, "0"))
        End If
    End If
    
    GaranteUnicoSepDecimal = Valor
    
    Exit Function
    
Erro_Overflow:
    GaranteUnicoSepDecimal = "0"

End Function

Public Function MaxCurrency(fValor, _
                            Optional fUsaAgrupDigitos As Boolean = True, _
                            Optional fPrecisao As Byte = 2, _
                            Optional fMaxVal As Currency = 922337203685477#) As Boolean
    
    If fValor = "" Then
        MaxCurrency = False
        Exit Function
    End If
    
    If Abs(fValor) > fMaxVal Then
        MaxCurrency = True
        
        MsgBox "O valor máximo permitido para este campo é: " & fMaxVal, vbInformation
        
        If fValor < 0 Then
            fMaxVal = -fMaxVal
        End If
        
        If fPrecisao = 0 Then
            fValor = FormataNumero(fMaxVal)
        Else
            fValor = FormataNumeroDecimal(fMaxVal, fUsaAgrupDigitos, False, fPrecisao)
        End If
    Else
        MaxCurrency = False
    End If
    
End Function

Public Function RemoveAgrupDigitos(fValor)
Dim i As Integer
Dim strValor As String, caract As String
Dim Negativo As String

    RemoveAgrupDigitos = ""
    
    strValor = CStr(fValor)
    
    If Left$(strValor, 1) = "-" Then
        Negativo = "-"
    Else
        Negativo = ""
    End If
    
    For i = 1 To Len(strValor)
        caract = Mid$(strValor, i, 1)
        
        If Not IsNumeric(caract) Then caract = ""
        
        RemoveAgrupDigitos = RemoveAgrupDigitos & caract
    Next
    
    RemoveAgrupDigitos = Negativo & RemoveAgrupDigitos

End Function

Public Function RemoveMoedaNumero(fValor)
Dim SimboloMoeda As String

    SimboloMoeda = VerificaSimboloMoeda()
        
    If Left$(fValor, Len(SimboloMoeda)) = SimboloMoeda Then
        RemoveMoedaNumero = Mid(fValor, Len(SimboloMoeda) + 1)
    Else
        RemoveMoedaNumero = fValor
    End If

End Function

'******************************************************************************
'* Função : TeclaNumero                                                       *
'* Descrição :                                                                *
'*      Verifica se um determinado código ANSI é numérico                     *
'*      Devolve o próprio código caso positivo, devolve 0 caso negativo       *
'*      o que impede que a tecla inválida digitada seja escrita no TextBox    *
'*      Caso o flag Permite_Decimal esteja ativo, permite que seja            *
'*      teclada o separador decimal definido no painel de controle            *
'*      Deve ser utilizado no evento KeyPress de um TextBox                   *
'*     Sintaxe no evento: KeyAscii = TeclaNumero(KeyAscii,True|False)         *
'* Parametros :                                                               *
'*      fKeyAscii - Código ANSI da tecla digitada                             *
'*      fPermite_Decimal - Flag (True ou Falso) que permite ou não que seja   *
'*                        teclada uma vírgula                                 *
'* Saída :                                                                    *
'*      O próprio código caso positivo, 0 caso negativo                       *
'******************************************************************************
Public Function TeclaNumero(fKeyAscii As Integer, _
                            Optional fPermite_Decimal As Boolean = False, _
                            Optional fPermite_Negativo As Boolean = False) As Integer
Dim KeyOut As Integer

    KeyOut = fKeyAscii

    If fKeyAscii <> vbKeyBack Then
        If fKeyAscii < vbKey0 Or fKeyAscii > vbKey9 Then
            If (fPermite_Decimal And Chr(fKeyAscii) = VerificaSepDecimalNumero()) Or _
               (fPermite_Negativo And (fKeyAscii = 45)) Then
            Else
                KeyOut = vbNull
            End If
        End If
    End If
    
    TeclaNumero = KeyOut

End Function

Public Function TestaSimboloMoeda(fSimboloMoeda As String) As Boolean
    
    If Trim$(UCase$(fSimboloMoeda)) = Trim$(UCase$(VerificaSimboloMoeda())) Then
        TestaSimboloMoeda = True
    Else
        TestaSimboloMoeda = False
        
        MsgBox "O Símbolo de Moeda deve corresponder ao formato cadastrado nas Configurações Regionais do Painel de Controle.", vbCritical
        
        On Error Resume Next
        Shell ("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,," & 2)
        Err = 0
    End If

End Function

Public Function TrocaSeparador(fValor, _
                               Optional fFormatoNumerico As gTpNumFormat = BR_NumFormat, _
                               Optional fRemoveAgrupDigitos As Boolean = False)
Dim i As Integer
Dim strValor As String, caract As String
Dim AchouSepDecimal As Boolean

    'Troca o formato dos separadores de acordo com a escolha de fFormatoNumerico
    'O 1º separador da direita para a esquerda será considerado como o separador Decimal

    TrocaSeparador = ""

    strValor = CStr(fValor)

    AchouSepDecimal = False

    For i = Len(strValor) To 1 Step -1
        caract = Mid$(strValor, i, 1)

        If caract = "," Or caract = "." Then
            If AchouSepDecimal Then
                If fRemoveAgrupDigitos Then
                    caract = ""
                Else
                    Select Case fFormatoNumerico
                        Case 0
                            caract = ","
                        Case 1
                            caract = "."
                    End Select
                End If
            Else
                AchouSepDecimal = True

                Select Case fFormatoNumerico
                    Case 0
                        caract = "."
                    Case 1
                        caract = ","
                End Select
            End If
        End If

        TrocaSeparador = TrocaSeparador & caract
    Next

    strValor = TrocaSeparador

    TrocaSeparador = ""

    For i = Len(strValor) To 1 Step -1
        caract = Mid$(strValor, i, 1)

        TrocaSeparador = TrocaSeparador & caract
    Next

End Function

'Public Function GetCurrency() As String
'Dim Symbol As String
'Dim iRet1 As Long
'Dim iRet2 As Long
'Dim lpLCDataVar As String
'Dim pos As Integer
'Dim Locale As Long
'
'   Locale = GetUserDefaultLCID()
'   iRet1 = GetLocaleInfo(Locale, LOCALE_SCURRENCY, lpLCDataVar, 0)
'
'   Symbol = String$(iRet1, 0)
'   iRet2 = GetLocaleInfo(Locale, LOCALE_SCURRENCY, Symbol, iRet1)
'   pos = InStr(Symbol, Chr$(0))
'   If pos > 0 Then
'      Symbol = Left$(Symbol, pos - 1)
'   End If
'
'   GetCurrency = Symbol
'
'End Function
'
Public Function VerificaSimboloMoeda() As String
Dim buffer As String * 100
Dim dl&

    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SCURRENCY, buffer, 99)
    VerificaSimboloMoeda = LPSTRtoVBString(buffer)
    
End Function

'******************************************************************************
'* Função   : VerificaFormatoMoeda                                            *
'* Descrição: Obtém o formato definido no Control Panel                       *
'*            através da função GetLocaleInfo                                 *
'* Saída    : formato de moeda do painel de controle                          *
'******************************************************************************
Public Function VerificaFormatoMoeda() As String

    VerificaFormatoMoeda = VerificaSepMilharMoeda + VerificaSepDecimalMoeda

End Function

Public Function VerificaSepMilharMoeda() As String
Dim buffer As String * 100
Dim dl&

    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHOUSANDSEP, buffer, 99)
    VerificaSepMilharMoeda = Trim$(LPSTRtoVBString(buffer))
    
End Function

Public Function VerificaSepDecimalMoeda() As String
Dim buffer As String * 100
Dim dl&

    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONDECIMALSEP, buffer, 99)
    VerificaSepDecimalMoeda = Trim$(LPSTRtoVBString(buffer))
    
End Function

'******************************************************************************
'* Função   : VerificaFormatoNumero                                           *
'* Descrição: Obtém o formato definido no Control Panel                       *
'*            através da função GetLocaleInfo                                 *
'* Saída    : formato de numero do painel de controle                         *
'******************************************************************************
Public Function VerificaFormatoNumero() As String

    VerificaFormatoNumero = VerificaSepMilharNumero + VerificaSepDecimalNumero
    
End Function

Public Function VerificaSepMilharNumero() As String
Dim buffer As String * 100
Dim dl&

    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, buffer, 99)
    VerificaSepMilharNumero = Trim$(LPSTRtoVBString(buffer))
    
End Function

Public Function VerificaSepDecimalNumero() As String
Dim buffer As String * 100
Dim dl&

    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, buffer, 99)
    VerificaSepDecimalNumero = Trim$(LPSTRtoVBString(buffer))
    
End Function

Public Function VerificaSimboloNegativo() As String
Dim buffer As String * 100
Dim dl&

    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SNEGATIVESIGN, buffer, 99)
    VerificaSimboloNegativo = Trim$(LPSTRtoVBString(buffer))
    
End Function

'******************************************************************************
'* Função : VerificaFormatoNumericoBRUS                                       *                       *
'* Descrição :                                                                *
'* Obtém o formato definido no Control Panel                                  *
'* Se o formato é "." para separador de milhares                              *
'*       e "," para separador decimal então retorna 0 senão 1                 *
'* Saída :                                                                    *
'*    BR_NumFormat - formato numerico compativel com o sistema (brasileiro)   *
'*    US_NumFormat - formato numerico incompatível com o sistema (americano)  *
'******************************************************************************
Public Function VerificaFormatoNumericoBRUS() As Byte
        
    If VerificaFormatoNumero() = ".," Then
        VerificaFormatoNumericoBRUS = BR_NumFormat
    Else
        VerificaFormatoNumericoBRUS = US_NumFormat
    End If

End Function

'******************************************************************************
'* Função : Centena                                                           *
'* Descrição : retorna a descrição de um numero                               *
'* Parametros :                                                               *
'*      Num - numero a ser descrito                                           *
'* Saída :                                                                    *
'*      descrição do numero definido                                          *
'******************************************************************************
Public Function Centena(Num)
Dim vet_alg(9)
Dim vet_onze(9)
Dim vet_dez(9)
Dim vet_cent(9)
Dim algarismo(3)
Dim i

    'FALTANDO TERMINAR
    
    vet_alg(1) = "um"
    vet_alg(2) = "dois"
    vet_alg(3) = "três"
    vet_alg(4) = "quatro"
    vet_alg(5) = "cinco"
    vet_alg(6) = "seis"
    vet_alg(7) = "sete"
    vet_alg(8) = "oito"
    vet_alg(9) = "nove"
            
    vet_onze(1) = "onze"
    vet_onze(2) = "doze"
    vet_onze(3) = "treze"
    vet_onze(4) = "quatorze"
    vet_onze(5) = "quinze"
    vet_onze(6) = "dezesseis"
    vet_onze(7) = "dezesete"
    vet_onze(8) = "dezoito"
    vet_onze(9) = "dezenove"
    
    vet_dez(1) = "dez"
    vet_dez(2) = "vinte"
    vet_dez(3) = "trinta"
    vet_dez(4) = "quarenta"
    vet_dez(5) = "cinquenta"
    vet_dez(6) = "sessenta"
    vet_dez(7) = "setenta"
    vet_dez(8) = "oitenta"
    vet_dez(9) = "noventa"
    
    For i = 1 To 9
    '    vet_dez(i) = vet_dez(i) + IIF(MOD(num, 10) = 0, "", " e")
    Next
    
    vet_cent(1) = IIf(Num = 100, "cem", "cento")
    vet_cent(2) = "duzentos"
    vet_cent(3) = "trezentos"
    vet_cent(4) = "quatrocentos"
    vet_cent(5) = "quinhentos"
    vet_cent(6) = "seiscentos"
    vet_cent(7) = "setecentos"
    vet_cent(8) = "oitocentos"
    vet_cent(9) = "novecentos"
    
    For i = 1 To 9
         vet_cent(i) = vet_cent(i) + IIf(Num = i * 100, "", " e")
    Next
    
    For i = 1 To 3
    '     algarismo(I) = Val(Mid$(Str(Num, 3), I, 1))
    Next
    
    Centena = ""
    
    For i = 1 To 3
        If algarismo(i) = 0 Then
    '        Loop
        End If
        If i = 2 And algarismo(2) = 1 And algarismo(3) > 0 Then
            Centena = Centena + Space(1) + vet_onze(algarismo(3))
            Exit For
        End If
        Select Case i
            Case 1
                Centena = Centena + Space(1) + vet_cent(algarismo(i))
            Case 2
                Centena = Centena + Space(1) + vet_dez(algarismo(i))
            Case 3
                Centena = Centena + Space(1) + vet_alg(algarismo(i))
        End Select
    Next

End Function

'******************************************************************************
'* Função : Extenso                                                           *
'* Descrição : retorna um numero por extenso                                  *
'* Parametros :                                                               *
'*      Num - numero a ser descrito                                           *
'* Saída :                                                                    *
'*      valor por extenso do numero definido                                  *
'******************************************************************************
Public Function Extenso(Num) As String
Dim vet_centena(3)
Dim Inteiro, i, CInteiro

    'FALTANDO TERMINAR
    
    Inteiro = Int(Num)
    'CInteiro = Trim$(STRZERO(Inteiro, 9))
    Extenso = ""
    
    For i = 1 To 3
        vet_centena(4 - i) = Val(Mid$(CInteiro, Len(CInteiro) - i * 3 + 1, 3))
    Next
    
    For i = 1 To 3
        If vet_centena(i) = 0 Then
            'Loop
        End If
    
        Extenso = Extenso + Centena(vet_centena(i))
    
        Select Case i
                Case 3
                     Extenso = Extenso
                Case 2
    '                 Extenso = Extenso + IIF(MOD(num, 1000) = 0, " mil", " mil e")
                Case 1
                     Extenso = Extenso + IIf(Num >= 2000000, " milhões", " milhão")
    '                 Extenso = Extenso + IIF(MOD(num,1000000) = 0, "", " e")
        End Select
    Next
    
    Extenso = Trim$(Extenso)
    
    If Left$(Extenso, 1) = "u" Then
        Extenso = "h" & Extenso
    End If
    
End Function

'******************************************************************************
'* Função : Unidade                                                           *
'* Descrição : retorna o valor por extenso de um numero no formato moeda      *
'* Parametros :                                                               *
'*      Valor - valor do numero a ser descrito                                *
'*      CUnidade - moeda utilizada                                            *
'* Saída :                                                                    *
'*      valor por extenso do numero definido                                  *
'******************************************************************************
Public Function unidade(fValor, CUnidade)
Dim plural, fracao, fracaop, muitos, centavos
    
    If fValor = 0 Then
        unidade = " "
    End If
    
    CUnidade = LCase(Trim$(CUnidade))
    
    Select Case CUnidade
            Case "dólar"
                    plural = "dólares"
                    fracao = "cent"
                    fracaop = "cents"
            Case "real"
                    plural = "reais"
                    fracao = "centavo"
                    fracaop = "centavos"
    End Select
    
    unidade = ""
    
    If Int(fValor) > 0 Then
            muitos = IIf(Int(fValor) > 1, True, False)
            unidade = Extenso(Int(fValor)) + IIf(muitos, " " & plural, " " & CUnidade)
    End If
    
    centavos = 100 * (fValor - Int(fValor))
    
    If centavos > 0 Then
            unidade = unidade + IIf(IsEmpty(unidade), "", " e") + Centena(centavos) & " " & IIf(centavos = 1, fracao, fracaop)
    End If

End Function

