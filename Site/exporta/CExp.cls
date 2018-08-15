VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Function exporta(arquivo As String, mes As Integer) As Boolean

   
Dim nomeWkSheet As String

Dim xlApp   As Excel.Application
Dim xlSheet As Excel.Worksheet
Dim GDB     As New ADODB.Connection
Dim Rs      As New ADODB.Recordset
Dim SQL     As String
Dim vl      As String
Dim desp    As String

If Not OpenExcelAppInstance(xlApp) Then
   exporta = False
   Exit Function
End If
    
arqXLS = "e:\fernando\projs\pessoal\grana\doc\contas.xls"
arqUDL = "exp.udl"
    
xlApp.Workbooks.Open arqXLS


Set xlSheet = xlApp.ActiveWorkbook.Sheets(3)

GDB.Open "File Name=" & App.Path & "\" & arqUDL

mes = 5

SQL = "SELECT D.nomeDespesa, Sum(G.valor) as ValTotal, count(*) as Qtd, " & _
   "ano, mes, FORMAT(diaPadrao, '00') as dd, receita " & _
   "FROM gasto AS G, despesa AS D " & _
   "WHERE G.idDespesa=D.idDespesa " & _
   "AND D.especial = true " & _
   "AND mes = " & mes & _
   " GROUP BY D.nomeDespesa, FORMAT(diaPadrao, '00'), ano, mes, " & _
   " receita " & _
   " UNION SELECT C.nomeCategoria, Sum(G.valor) as ValTotal, count(*) as Qtd, " & _
      "ano, mes, '99' as dd, IIF(C.nomeCategoria = 'RECEITA', TRUE, false) " & _
      "FROM gasto AS G, despesa AS D, categoria AS C " & _
      "WHERE G.idDespesa=D.idDespesa " & _
      "AND D.idCategoria=C.idCategoria " & _
      "AND D.especial = false " & _
      "AND mes = " & mes & _
      " GROUP BY  C.nomeCategoria, ano, mes ORDER BY dd"



Rs.Open SQL, GDB

While Not Rs.EOF

   vl = Replace(Rs(1), ",", ".")
   desp = UCase(Trim(Rs(0)))
   If CInt(Rs(5)) <> 99 And CInt(Rs(5)) <> 0 Then
      desp = Rs(5) & " - " & desp
   End If
   
   For i = 3 To 52
      
      If desp = UCase(Trim(xlSheet.Cells(i, 1))) Then
         If Rs(6) = True Then
            xlSheet.Cells(i, mes + 1) = vl
         Else
            xlSheet.Cells(i, mes + 1) = "-" & vl
         End If
         Exit For
      End If
   
   Next
  Rs.MoveNext

Wend

Rs.Close

GDB.Close

xlApp.ActiveWorkbook.Save

xlApp.ActiveWorkbook.Close False
    
CloseExcelAppInstance xlApp


End

End Function


