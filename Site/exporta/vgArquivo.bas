Attribute VB_Name = "vgArquivo"
Option Explicit

'----------------------------------------
' Declaração de APIs e Constantes Locais
'----------------------------------------
'LeIni()
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
         ByVal lpKeyName As Any, _
         ByVal lpDefault As String, _
         ByVal lpReturnedString As String, _
         ByVal nSize As Long, _
         ByVal lpFileName As String) As Long

'EscreveIni()
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
         ByVal lpKeyName As Any, _
         ByVal lpString As Any, _
         ByVal lpFileName As String) As Long

'WinDir()
Private Const gintMAX_PATH_LEN As Integer = 260

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
        (ByVal lpBuffer As String, _
         ByVal nSize As Long) As Long

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
        (ByVal lpBuffer As String, _
         ByVal nSize As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
        (ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" _
        (ByVal lpszPath As String, _
         ByVal lpPrefixString As String, _
         ByVal wUnique As Long, _
         ByVal lpTempFileName As String) As Long

'FileErrors()
Private Const mnErrDeviceUnavailable = 68
Private Const mnErrDiskNotReady = 71
Private Const mnErrDeviceIO = 57
Private Const mnErrDiskFull = 61
Private Const mnErrBadFileName = 64
Private Const mnErrBadFileNameOrNumber = 52
Private Const mnErrPathDoesNotExist = 76
Private Const mnErrBadFileMode = 54
Private Const mnErrFileAlreadyOpen = 55
Private Const mnErrPermissionDenied = 70
Private Const mnErrInputPastEndOfFile = 62

'---------------------------------
' Declaração de Constantes Globais
'---------------------------------
'AbreArquivo()
Public Enum tpOpenFile
    tpOFAppend = 1
    tpOFOutput = 2
    tpOFInput = 3
End Enum

Public Enum tpLockFile
    tpLFShared = 1
    tpLFLockRead = 2
    tpLFLockWrite = 3
    tpLFLockReadWrite = 4
End Enum

Public Function AbreArquivo(fArquivo As String, _
                            fTipoAbertura As tpOpenFile, _
                            Optional fTipoLock As tpLockFile = tpLFShared) As Integer
Dim FileNumber As Integer

    On Error GoTo erro_arq
    
    AbreArquivo = 0
    
    FileNumber = FreeFile
    
    Select Case fTipoAbertura
    Case 1
        Select Case fTipoLock
        Case 1
            Open fArquivo For Append Shared As #FileNumber
        Case 2
            Open fArquivo For Append Lock Read As #FileNumber
        Case 3
            Open fArquivo For Append Lock Write As #FileNumber
        Case 4
            Open fArquivo For Append Lock Read Write As #FileNumber
        End Select
    Case 2
        Select Case fTipoLock
        Case 1
            Open fArquivo For Output Shared As #FileNumber
        Case 2
            Open fArquivo For Output Lock Read As #FileNumber
        Case 3
            Open fArquivo For Output Lock Write As #FileNumber
        Case 4
            Open fArquivo For Output Lock Read Write As #FileNumber
        End Select
    Case 3
        Select Case fTipoLock
        Case 1
            Open fArquivo For Input Shared As #FileNumber
        Case 2
            Open fArquivo For Input Lock Read As #FileNumber
        Case 3
            Open fArquivo For Input Lock Write As #FileNumber
        Case 4
            Open fArquivo For Input Lock Read Write As #FileNumber
        End Select
    End Select
    
    AbreArquivo = FileNumber
    
    Exit Function
    
erro_arq:
    TrataErro_Arq fArquivo

End Function

'Public Function AbreDialogoSalvaArqTexto(CDialog As CommonDialog, _
                                         Optional fDialogTitle As String = "", _
                                         Optional fFileName As String = "", _
                                         Optional fFilter As String = "Arquivos Texto (*.txt)|*.txt", _
                                         Optional fDefaultExt As String = "txt", _
                                         Optional fOverwritePrompt As Boolean = True) As String

    
'    AbreDialogoSalvaArqTexto = ""
    
'    On Error GoTo ErrHandler
    
'    With CDialog
    
'    .CancelError = True
    
'    If fOverwritePrompt Then
'        .FLAGS = cdlOFNOverwritePrompt
'    Else
'        .FLAGS = 0
'    End If
    
'    .DialogTitle = fDialogTitle
    
'    .FileName = fFileName
    
'    .Filter = fFilter
'    .DefaultExt = fDefaultExt
    
'    .ShowSave
    
'    AbreDialogoSalvaArqTexto = .FileName
    
'    End With
    
'ErrHandler:
  'User pressed the Cancel button
'  Exit Function

'End Function

Public Function AbreviaDir(fDir As String, _
                           fTamanho As Integer) As String
Dim Pos1stDir As Integer
Dim DirLeft As String

    If fTamanho <= 4 Then
        AbreviaDir = Left$(fDir, fTamanho)
        Exit Function
    End If
    
    If Left$(fDir, 2) = "\\" Then
        'ex: "\\Dados"
        Pos1stDir = 3
    Else
        'ex: "C:\"
        Pos1stDir = 4
    End If
    
    Pos1stDir = InStr(Pos1stDir, fDir, "\")
    
    'String até o primeiro diretório
    DirLeft = Left$(fDir, Pos1stDir) & "..."
    
    If fTamanho <= Len(DirLeft) Then
        AbreviaDir = Left$(fDir, fTamanho)
    Else
        AbreviaDir = DirLeft & Right$(fDir, fTamanho - Len(DirLeft))
    End If

End Function

Public Function AlinhaNumero(fValor As Variant, _
                             Optional fTamanho As Integer = 0, _
                             Optional RetiraDecimal As Boolean = False) As String
Dim Valor As Variant

    If RetiraDecimal Then
        Valor = CDbl(vgNumero.FormataNumeroDecimal(fValor * 10))
    Else
        Valor = fValor * 10
    End If
    
    If fTamanho = 0 Then
        fTamanho = Len(fValor)
    End If
    
    AlinhaNumero = Right$(String(fTamanho, "0") & Trim$(Left$(Valor, fTamanho)), fTamanho)

End Function

Public Function AlinhaTexto(fValor As Variant, _
                            Optional fTamanho As Integer = 0) As String

    If fTamanho = 0 Then
        fTamanho = Len(fValor)
    End If
    
    AlinhaTexto = Left$(Trim$(Left$(fValor, fTamanho)) & String(fTamanho, Space(1)), fTamanho)

End Function

'******************************************************************************
'* Função : BotaBarraDir                                                      *
'* Descrição : coloca uma barra no final de um string de diretório            *
'* Parametros :                                                               *
'*      fDiretorio - diretório a alterar                                      *
'* Saída :                                                                    *
'*      Diretório com a barra '\'                                             *
'******************************************************************************
Public Function BotaBarraDir(fDiretorio As String) As String
    
    If Right$(fDiretorio, 1) <> "\" Then
        fDiretorio = fDiretorio & "\"
    End If
    
    BotaBarraDir = fDiretorio

End Function

Public Function CheckFNLength(fStrFilename As String) As Boolean
' This routine verifies that the length of the filename fStrFilename is valid.
' Under NT (Intel) and Win95 it can be up to 259 (gintMAX_PATH_LEN-1) characters
' long.  This length must include the drive, path, filename, commandline
' arguments and quotes (if the string is quoted).
    
    CheckFNLength = (Len(fStrFilename) < gintMAX_PATH_LEN - 1)
    
End Function

Public Sub CreateLongDir(sDrive As String, _
                         sDir As String)
Dim sBuild As String
    
    While InStr(2, sDir, "\") > 1
        sBuild = sBuild & Left$(sDir, InStr(2, sDir, "\") - 1)
        
        sDir = Mid$(sDir, InStr(2, sDir, "\"))
        
        If Dir$(sDrive & sBuild, vbDirectory) = "" Then
            MkDir sDrive & sBuild
        End If
    Wend

End Sub

'******************************************************************************
'* Função : EscreveIni                                                        *
'* Descrição : altera uma clausula em um Arquivo ini                          *
'* Parametros :                                                               *
'*      fArquivo - nome do Arquivo ini                                        *
'*      fSecao - Secao a alterar do Arquivo ini                               *
'*      fChave - clausula a alterar                                           *
'*      fValor - novo Valor                                                   *
'* Saída :                                                                    *
'*      true - operação concluída com sucesso                                 *
'*      false - operação não concluída                                        *
'******************************************************************************
Public Function EscreveIni(fArquivo As String, _
                           fSecao As String, _
                           fChave As String, _
                           fValor As String) As Boolean
Dim NumArq As Integer
Dim Tamanho As Long
Dim Resp As Long
    
    On Error GoTo err_EscreveIni
    
    EscreveIni = True
    
    NumArq = FreeFile
    
    If Dir(fArquivo) = "" Then
        'Tenta abrir um novo arquivo caso não exista
        Open fArquivo For Append Shared As #NumArq
        Close #NumArq
    End If

    Resp = WritePrivateProfileString(fSecao, fChave, fValor, fArquivo)

    Exit Function

err_EscreveIni:
    TrataErro_Arq fArquivo
    EscreveIni = False
    Exit Function

End Function

'******************************************************************************
'* Função : ExisteArquivo                                                     *
'* Descrição :  Verifica a existência de um Arquivo.                          *
'* Retorna Verdadeiro se existe o Arquivo ou Falso caso contrário.            *
'* Parametros :                                                               *
'*      fArquivo - String com o caminho do Arquivo a verificar                *
'* Saída :                                                                    *
'*      true - o Arquivo existe                                               *
'*      false - o Arquivo não existe                                          *
'******************************************************************************
Public Function ExisteArquivo(fArquivo As String) As Boolean
Dim dummy As String

    On Error Resume Next
    
    dummy = ""
    dummy = Dir$(fArquivo, vbNormal)

    If dummy <> "" Then
        ExisteArquivo = True
    Else
        ExisteArquivo = False
    End If

    Err = 0

End Function

'******************************************************************************
'* Função : ExisteDir                                                         *
'* Descrição : verifica se existe o diretório especificado                    *
'* Parametros :                                                               *
'*      fPath - Path a verificar                                              *
'* Saída  :                                                                   *
'*      true - O Diretório Existe                                             *
'*      false - O Diretório não existe                                        *
'******************************************************************************
Public Function ExisteDir(fPath As String) As Boolean
Dim dummy As String

    On Error Resume Next

    fPath = BotaBarraDir(fPath)
    
    dummy = Dir$(fPath, vbDirectory)
    If dummy = "" Then
        ExisteDir = False
    Else
        ExisteDir = True
    End If

    Err = 0
    
End Function

'-----------------------------------------------------------
' FUNCTION: Extension
'
' Extracts the extension portion of a file/path name
'
' IN: [fStrFilename] - file/path to get the extension of
'
' Returns: The extension if one exists, else ""
'-----------------------------------------------------------
Public Function Extension(fStrFilename As String) As String
Dim intPos As Integer

    Extension = ""

    intPos = Len(fStrFilename)

    Do While intPos > 0
        Select Case Mid$(fStrFilename, intPos, 1)
            Case "."
                Extension = Mid$(fStrFilename, intPos + 1)
                Exit Do
            Case "/", "\"
                Exit Do
        End Select

        intPos = intPos - 1
    Loop

End Function

Public Sub FechaArquivo(fArqNum As Integer)

    Close fArqNum
    
End Sub

'******************************************************************************
'Erros de manipulação de arquivo
' Return Value      Meaning
' 0                 Resume
' 1                 Resume Next
' 2                 Unrecoverable error
' 3                 Unrecognized error
'******************************************************************************
Public Function FileErrors() As Integer
Dim intMsgType As Integer
Dim strMsg As String
Dim intResponse As Integer
    
    Screen.MousePointer = vbDefault
    
    intMsgType = vbExclamation
    
    Select Case Err.Number
        Case mnErrDeviceUnavailable             ' Error 68
            'strMsg = "That device appears unavailable."
            strMsg = "Este dispositivo parece estar indisponível."
            intMsgType = vbExclamation + vbOKCancel
        Case mnErrDiskNotReady                  ' Error 71
            'strMsg = "Insert a disk in the drive and close the door."
            strMsg = "Insira um disco no drive e feche a porta."
            intMsgType = vbExclamation + vbOKCancel
        Case mnErrDeviceIO                      ' Error 57
'            strMsg = "Internal disk error."
            strMsg = "Erro interno de disco."
            intMsgType = vbExclamation + vbOKOnly
        Case mnErrDiskFull                      ' Error 61
'            strMsg = "Disk is full. Continue?"
            strMsg = "O disco está cheio. Continua ?"
            intMsgType = vbExclamation + vbAbortRetryIgnore
        Case mnErrBadFileName, mnErrBadFileNameOrNumber ' Error 64 & 52
'            strMsg = "That filename is illegal."
            strMsg = "Este nome de arquivo é inválido."
            intMsgType = vbExclamation + vbOKCancel
        Case mnErrPathDoesNotExist                ' Error 76
'            strMsg = "That path doesn't exist."
            strMsg = "Este caminho não existe."
            intMsgType = vbExclamation + vbOKCancel
        Case mnErrBadFileMode                     ' Error 54
'            strMsg = "Can't open your file for that type of access."
            strMsg = "Não é possível abrir o arquivo para este tipo de acesso."
            intMsgType = vbExclamation + vbOKOnly
        Case mnErrFileAlreadyOpen             ' Error 55
'            strMsg = "This file is already open."
            strMsg = "Este arquivo já está aberto."
            intMsgType = vbExclamation + vbOKOnly
        Case mnErrInputPastEndOfFile              ' Error 62
'            strMsg = "This file has a nonstandard end-of-file marker, "
'            strMsg = strMsg & "or an attempt was made to read beyond "
'            strMsg = strMsg & "the end-of-file marker."
            strMsg = "Este arquivo tem um marcador de fim de arquivo não-padrão, "
            strMsg = strMsg & "ou foi feita uma tentativa para ler além "
            strMsg = strMsg & "do marcador de fim de arquivo."
            intMsgType = vbExclamation + vbAbortRetryIgnore
        Case mnErrPermissionDenied                ' Error 70
            strMsg = "Permissão negada ao tentar abrir este arquivo."
            intMsgType = vbExclamation + vbOKOnly
        Case Else:
            strMsg = "Error " & Str(Err.Number) & ": " & Err.Description
            intMsgType = vbCritical + vbOKOnly
            MsgBox strMsg, intMsgType, "Erro de Disco"
            FileErrors = 3
            Exit Function
    End Select
    
    intResponse = MsgBox(strMsg, intMsgType, "Erro de Disco")
    
    Select Case intResponse
        Case 1, 4       ' OK, Retry buttons.
            FileErrors = 0
        Case 2, 5       ' Cancel, Ignore buttons.
            FileErrors = 1
        Case 3          ' Abort button.
            FileErrors = 2
    End Select

End Function

'******************************************************************************
'* Função :   LeIni                                                           *
'* Descrição : le uma clausula de um Arquivo ini                              *
'* Parametros :                                                               *
'*      fArquivo - Arquivo ini a ser lido                                     *
'*      fSecao - Secao do Arquivo ini                                         *
'*      fChave - clausula a ser lida                                          *
'* Saída :                                                                    *
'*      conteúdo da clausula definida                                         *
'******************************************************************************
Public Function LeIni(fArquivo As String, _
                      fSecao As String, _
                      fChave As String) As String
Dim CharNaoAchou As String
Dim Tamanho As Long
Dim Resp As Long
Dim StringRetorno As String

    CharNaoAchou = "*"
    StringRetorno = Space$(512)
    Tamanho = Len(StringRetorno)

    Resp = GetPrivateProfileString(fSecao, fChave, CharNaoAchou, StringRetorno, Tamanho, fArquivo)
    LeIni = Left$(StringRetorno, Resp)

End Function

'******************************************************************************
'* Função :  MudaDir                                                          *
'* Descrição : muda o diretório corrente                                      *
'* Parametros :                                                               *
'*      fDiretorio - novo diretório corrente                                  *
'* Saída :                                                                    *
'*      true - operação realizada com sucesso                                 *
'*      false - operação não realizada                                        *
'******************************************************************************
Public Function MudaDir(fDiretorio As String) As Boolean
Dim SaiRotina As Boolean
    
    MudaDir = False
    SaiRotina = False

    Select Case Len(fDiretorio)
        Case 1
           Exit Function
        Case 2
            fDiretorio = BotaBarraDir(fDiretorio)
        Case Is > 3
            fDiretorio = TiraBarraDir(fDiretorio)
    End Select
    
    On Error GoTo Erro

    ChDrive Left$(fDiretorio, 2)
    
    If SaiRotina Then Exit Function

    ChDir fDiretorio
    
    If SaiRotina Then Exit Function
    
    If CurDir = fDiretorio Then
        MudaDir = True
    End If

    Exit Function

Erro:
    SaiRotina = True
    Resume Next

End Function

'******************************************************************************
'* Função : TiraBarraDir                                                      *
'* Descrição : retira a ultima barra na definição de um diretório             *
'* Parametros :                                                               *
'*      fDiretorio - Diretório a alterar                                      *
'* Saída :                                                                    *
'*      string sem a barra '\'                                                *
'******************************************************************************
'Public Function TiraBarraDir(fDiretorio As String) As String
    
'    If Right$(fDiretorio, 1) = "\" Then
'         fDiretorio = Left$(fDiretorio, Len(fDiretorio) - 1)
'    End If

'    TiraBarraDir = fDiretorio

'End Function

'******************************************************************************
'* Função : VejoDir                                                           *
'* Descrição : verifica se existe um diretório e muda o diretório corrente    *
'* Parametros :                                                               *
'*      fDiretorio - diretório a verificar                                    *
'* Saída :                                                                    *
'*      true - diretório localizado, diretório corrente alterado              *
'*      false - diretório não localizado                                      *
'******************************************************************************
Public Function VejoDir(fDiretorio As String) As Boolean
    
    fDiretorio = TiraBarraDir(UCase$(fDiretorio))
    
    If Len(fDiretorio) = 2 Then
        fDiretorio = BotaBarraDir(fDiretorio)
    End If

    VejoDir = False
    
    On Error Resume Next
    
    ChDrive Left$(fDiretorio, 1)
        
    ChDir fDiretorio
    
    If CurDir = fDiretorio Then
        VejoDir = True
    End If
    
    Err = 0

End Function

'******************************************************************************
'* Função : WinDir                                                            *
'* Descrição : retorna o diretório do windows                                 *
'* Saída : caminho do diretório do windows                                    *
'*                                                                            *
'******************************************************************************
Public Function WinDir() As String
Dim Tamanho As Long
Dim Resp As Long
Dim StringRetorno As String

    StringRetorno = Space$(512)
    Tamanho = Len(StringRetorno)

    Resp = GetWindowsDirectory(StringRetorno, Tamanho)
    WinDir = Left$(StringRetorno, Resp)

End Function

Public Function RemovePath(fNomeArq As String) As String
Dim PosLastDir As Integer

    PosLastDir = InStrRev(fNomeArq, "\")
    
    RemovePath = Mid$(fNomeArq, PosLastDir + 1)

End Function

Public Function SysDir() As String
Dim Tamanho As Long
Dim Resp As Long
Dim StringRetorno As String

    StringRetorno = Space$(512)
    Tamanho = Len(StringRetorno)

    Resp = GetSystemDirectory(StringRetorno, Tamanho)
    SysDir = Left$(StringRetorno, Resp)

End Function

Public Function TempDir() As String
Dim Tamanho As Long
Dim Resp As Long
Dim StringRetorno As String

    StringRetorno = Space$(512)
    Tamanho = Len(StringRetorno)

    Resp = GetTempPath(Tamanho, StringRetorno)
    TempDir = Left$(StringRetorno, Resp)

End Function

'-----------------------------------------------------------
' FUNCTION: TempFilename
' Get a temporary filename for a specified drive and
' filename prefix
' PARAMETERS:
'   fStrDestPath - Location where temporary file will be created.  If this
'                 is an empty string, then the location specified by the
'                 tmp or temp environment variable is used.
'   fLpPrefixString - First three characters of this string will be part of
'                    temporary file name returned.
'   fWUnique - Set to 0 to create unique filename.  Can also set to integer,
'             in which case temp file name is returned with that integer
'             as part of the name.
'   fLpTempFilename - Temporary file name is returned as this variable.
' RETURN:
'   True if function succeeds; false otherwise
'-----------------------------------------------------------
'Public Function TempFileName(ByVal fStrDestPath As String, _
'                             ByVal fLpPrefixString As String, _
'                             ByVal fWUnique As Long, _
'                             fLpTempFilename As String) As Boolean
'
'    If fStrDestPath = gstrNULL Then
'        ' No destination was specified, use the temp directory.
'        fStrDestPath = String(gintMAX_PATH_LEN, vbNullChar)
'        If GetTempPath(gintMAX_PATH_LEN, fStrDestPath) = 0 Then
'            GetTempFileName = False
'            Exit Function
'        End If
'    End If
'    fLpTempFilename = String(gintMAX_PATH_LEN, vbNullChar)
'    TempFileName = GetTempFileName(fStrDestPath, fLpPrefixString, fWUnique, fLpTempFilename) > 0
'    fLpTempFilename = StripTerminator(fLpTempFilename)
'
'End Function

Public Sub TrataErro_Arq(Optional fArquivo As String = "")
Dim Texto As String

    Screen.MousePointer = vbDefault
    
    Texto = "Erro no arquivo: " & fArquivo & vbCrLf
    
    Select Case Err
    Case 53:
        Texto = Texto & "Motivo: Arquivo não encontrado."
    
    Case 57, 68, 71:
        Texto = Texto & "Motivo: Erro de acesso ao Disco."
    
    Case 52, 54:
        Texto = Texto & "Motivo: Erro de acesso ao Arquivo."
    
    Case 75:
        Texto = Texto & "Motivo: Erro de acesso ao Diretório/Arquivo."
    
    Case 76:
        Texto = Texto & "Motivo: Caminho não encontrado."

    Case 61:
        Texto = Texto & "Motivo: Disco cheio."

    Case 67:
        Texto = Texto & "Motivo: Muitos arquivos abertos."

    Case 76:
        Texto = Texto & "Motivo: Caminho não encontrado."

    Case 55, 70:
        Texto = Texto & "Motivo: O Arquivo está aberto por outro aplicativo."
    
    Case Else:
        Texto = Texto & vbCrLf
        Texto = Texto & "Erro " & Str(Err.Number) & ": " & Err.Description
    End Select
   
    MsgBox Texto, vbCritical, "ERRO"

End Sub

'******************************************************************************
'* Função : TiraBarraDir                                                      *
'* Descrição : retira a ultima barra na definição de um diretório             *
'* Parametros :                                                               *
'*      fDiretorio - Diretório a alterar                                      *
'* Saída :                                                                    *
'*      string sem a barra '\'                                                *
'******************************************************************************
Public Function TiraBarraDir(fDiretorio As String) As String

    If Right$(fDiretorio, 1) = "\" Then
        fDiretorio = Left$(fDiretorio, Len(fDiretorio) - 1)
    End If

    TiraBarraDir = fDiretorio

End Function

