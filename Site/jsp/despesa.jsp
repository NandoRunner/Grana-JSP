<%@ page import="java.util.*" %>
<%@ page import="grana.*" %>

<html>
<head>
<title>Grana 2003 for Web - Cadastro de despesas e receitas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<%@ include file ="inc/validaCampo.js" %>

<script language="JavaScript">

function testaCampos(){
  
  if (!validaCampo(document.frm.nomeDespesa, "Despesa"))
  	return false;

  if (!validaCampo(document.frm.idCateg, "Categoria"))
  	return false;

	document.frmEditar.submit();	
  return true;	
}

</script>

<link href="grana.css" rel="stylesheet" type="text/css">
</head>

<jsp:useBean id="despesa" scope="session" class="grana.Despesa" />
<jsp:setProperty name="despesa" property="*" />

<body>
<a name="topo"></a> 
<table width="550" height="20" border="0" align="center" cellpadding="10" cellspacing="0">
  <tr> 
    <td height="22" align="center" valign="middle" bgcolor="#009999"><font color="#FFFFFF" size="4" face="Comic Sans MS">Cadastro 
      de despesas e receitas</font></td>
  </tr>
</table>
<br>
<form name="frm" action="despesa.jsp" method="post" onSubmit="return testaCampos()">
  <table width="500" height="169" border="0" align="center" cellspacing="5">
    <tr> 
      <td height="36" align="right" valign="middle" class="unnamed1"><font color="#009999">Despesa/receita</font></td>
      <td height="36" colspan="3" valign="middle" class="unnamed1">
	   <input name="nomeDespesa" value=<%=request.getParameter("nomeDespesa")%> type="text" id="nomeCategoria4" size="30" maxlength="30"></td>
    </tr>
    <tr> 
      <td height="36" align="right" valign="middle" class="unnamed1"><font color="#009999">Categoria</font></td>
      <td height="36" colspan="3" valign="middle" class="unnamed1"> <select name="idCateg" id="idCateg">
          <option value=""></option>
          <jsp:useBean id="categ" scope="session" class="grana.Categoria" />
          <jsp:setProperty name="categ" property="*" />
          <% 	String[] items = categ.montaLista();
	for (int i=0; i<items.length; i++) {
%>
          <option value="<%= items[i] %>"><%= items[i] %></option>
          <%	}
%>
        </select></td>
    </tr>
    <tr> 
      <td height="36" align="right"><font color="#009999">Valor padr&atilde;o</font></td>
      <td height="36"><input name="valorPadrao" type="text" id="valorPadrao2" size="15" maxlength="10"></td>
      <td><font color="#009999">Dia padr&atilde;o</font></td>
      <td><input name="diaPadrao" type="text" id="diaPadrao" size="5" maxlength="2"></td>
    </tr>
    <tr> 
      <td height="36" align="right"><font color="#009999">Especial</font></td>
      <td height="36" colspan="3"><input name="especial" type="checkbox" id="especial2" value="true"></td>
    </tr>
  </table>
  <br>
  <table width="220" border="0" align="center" cellspacing="5">
    <tr> 
      <td height="26"><div align="center"> 
          <input name="Submit" type="submit" value="salvar">
        </div></td>
      <td><div align="center"> 
          <input name="Submit2" type="reset" value="Limpar">
        </div></td>
    </tr>
  </table>
  <br>
  <table width="350" height="36" border="0" align="center" cellspacing="5">
    <tr align="left" valign="middle"> 
      <td height="36" colspan="2"> 
<%			if (despesa.existe(request)) {
%>
        <p>Despesa <font color="#3399FF"><%=request.getParameter("nomeDespesa") %> </font>já existe!</p>
<%			} else {
				if (despesa.insere(request)) { 
%>
        <p>Despesa <font color="#3399FF"><%=request.getParameter("nomeDespesa") %> </font>cadastrada!</p>
<%				}
			}
%>
    </tr>
  </table>
  <br>
</form>
<br>
<table width="350" border="0" align="center" bgcolor="#FFFFFF">
  <tr bgcolor="#009999"> 
    <td width="298" bgcolor="#FFFFFF"><div align="center"><font color="#333333" size="2" face="Comic Sans MS">Despesas/receitas 
        por categoria</font></div></td>
  </tr>

<%  Vector tab =  despesa.montaRelatorio();
  	String nomeCateg = "";		
	for (int i=0, j=0; i<tab.size(); i++, j++) {
		if (!(nomeCateg.equals(((TabDespesa)tab.elementAt(i)).getNomeCategoria()))){
 %>
 </table>  
<br><br>
<table width="280" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr bgcolor="#009999">
    <td width="180"><font color="#FFFFFF" size="2" face="Comic Sans MS"> <%=((TabDespesa)tab.elementAt(i)).getNomeCategoria() %></font></td>
    <td width="18"><div align="left"></div></td>
  </tr>
  <% 		
			j = 0;
			nomeCateg = ((TabDespesa)tab.elementAt(i)).getNomeCategoria();
		} else if (j % 2 == 0) {
%>
  <tr bgcolor="#ffffff"> 
    <%  	} else {
%>
  <tr bgcolor="#C2E9E7">
    <%		}
%>
    <td width="180"><font face="Comic Sans MS" size="2"><%=((TabDespesa)tab.elementAt(i)).getNome() %></font></td>
	<td width="18">
    <a href ="despesa.jsp?nomeDespesa=<%=((TabDespesa)tab.elementAt(i)).getNome()%>">
	<img src="img/editar.gif" width="13" height="12" border="0"></a></td>
  </tr>
  <%	}
%>
</table>

<%@ include file ="inc/base.inc" %>

</body>
</html>
