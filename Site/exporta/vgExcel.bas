Attribute VB_Name = "vgExcel"
Option Explicit

'----------------------------------
' Declaração de Constantes Globais
'----------------------------------
Public Const cMaxLenExcelBookName As Byte = 30

Public Function OpenExcelAppInstance(xlApp As Excel.Application) As Boolean
Dim XL_NOTRUNNING As Boolean 'Long = 429

    On Error GoTo ShowName_Err
    
    OpenExcelAppInstance = True
    
    XL_NOTRUNNING = True
    
    Set xlApp = GetObject(, "Excel.Application")
    
OpenNew:
    'Excel is not currently running (when called from here...)
    Set xlApp = New Excel.Application
    
    Exit Function

ShowName_Err:
    If XL_NOTRUNNING Then
        XL_NOTRUNNING = False
        
        Resume OpenNew
    Else
        Screen.MousePointer = vbDefault
        
        OpenExcelAppInstance = False
        
        TrataErro
    
        Set xlApp = Nothing
    End If

End Function

Public Sub CloseExcelAppInstance(xlApp As Excel.Application)

    On Error Resume Next
    
    xlApp.Quit
    
    Set xlApp = Nothing
    
End Sub

Public Function CreateNewWorkbook(xlApp As Excel.Application, _
                                  wkbNew As Excel.Workbook, _
                                  fBookName As String, _
                                  Optional fNumSheets As Integer = 3, _
                                  Optional fSaveAsDialog As Boolean = True, _
                                  Optional fCreateBackup As Boolean = False, _
                                  Optional fAddToMRU As Boolean = True) As Boolean
Dim OrigNumSheets As Integer

    ' This procedure creates a new workbook file and saves it by using the path
    ' and name specified in the fBookName argument. You use the fNumsheets
    ' argument to specify the number of worksheets in the workbook;
    ' the default is 3.
    
    CreateNewWorkbook = False
    
    On Error GoTo CreateNew_Err
    
    OrigNumSheets = xlApp.SheetsInNewWorkbook
    If OrigNumSheets <> fNumSheets Then
        xlApp.SheetsInNewWorkbook = fNumSheets
    End If
    
    Set wkbNew = xlApp.Workbooks.Add
    
    xlApp.SheetsInNewWorkbook = OrigNumSheets
   
    If fSaveAsDialog Then
        wkbNew.Activate
        
        'Faltando: Não consigo passar o filtro
        fBookName = wkbNew.Application.GetSaveAsFilename(fBookName & ".xls")
        ', "All Files (*.*)|*.*|Arquivos do Excel (*.xls)|*.xls", 2
    End If
    
    fBookName = Trim$(fBookName)
    
    If fBookName = "False" Then Exit Function
    
    If Right(fBookName, 1) = "." Then
        fBookName = fBookName & "xls"
    End If
       
    'fCreateBackup não está funcionando,
    'pois ao salvar o arquivo aberto após sair desta função
    'é criado um novo backup do arquivo vazio
    If fCreateBackup Then
        'Só cria backup se o arquivo já existe
        fCreateBackup = vgArquivo.ExisteArquivo(fBookName)
    End If
    
    wkbNew.SaveAs fBookName, , , , , fCreateBackup, xlExclusive, , fAddToMRU
    
    CreateNewWorkbook = True
    
    Exit Function

CreateNew_Err:
    Screen.MousePointer = vbDefault
    
    If InStr(UCase(Err.Description), "SAVEAS") = 0 Then
        TrataErro
    End If
    
    Set wkbNew = Nothing

End Function

Public Sub TrataErro(Optional ftxtErro As String)
Dim Texto As String
   
    Screen.MousePointer = vbDefault
    
    Texto = ""
    If ftxtErro <> "" Then
        Texto = "Erro no evento: " & ftxtErro & vbCrLf
    End If
    Texto = Texto & "Erro " & Str(Err.Number) & ": " & Err.Description
    
    MsgBox Texto, vbExclamation, "ERRO"
   
End Sub
