<html>
<head>
<title>Grana 2003 for Web - Cadastro de categorias</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<%@ include file ="inc/validaCampo.js" %>

<script language="JavaScript">

function testaCampos(){
  
  if (!validaCampo(document.frm.nomeCategoria, "Categoria"))
  	return false;
}
</script>

<link href="grana.css" rel="stylesheet" type="text/css">
</head>

<jsp:useBean id="categ" scope="session" class="grana.Categoria" />
<jsp:setProperty name="categ" property="*" />

<body>
<table width="550" height="20" border="0" align="center" cellpadding="10" cellspacing="0">
  <tr> 
    <td height="22" align="center" valign="middle" bgcolor="#009999"><font color="#FFFFFF" size="4" face="Comic Sans MS">Cadastro 
      de categorias </font></td>
  </tr>
</table>
<br>
<form name="frm" action="categoria.jsp" method="post" onSubmit="return testaCampos()">
  <table width="350" height="71" border="0" align="center" cellspacing="5">
    <tr valign="middle"> 
      <td width="57" height="36" align="right" class="unnamed1"><font color="#009999">Categoria</font></td>
      <td width="224" height="36" class="unnamed1"><input name="nomeCategoria" type="text" id="nomeCategoria" size="30" maxlength="30"></td>
    </tr>
    <tr align="left" valign="middle"> 
      <td height="36" colspan="2">
<%	if (categ.existe(request)) {
%>
		<p>Categoria <font color="#3399FF"><%=request.getParameter("nomeCategoria") %> </font>já existe!</p>
<%	} else {
		if (categ.insere(request)) { 
%>
			<p>Categoria <font color="#3399FF"><%=request.getParameter("nomeCategoria") %> </font>cadastrada!</p>
<%		}
	}
%>

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
  </form>
<br>
<table width="250" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr bgcolor="#009999"> 
    <td width="250"><div align="left"><font color="#FFFFFF" size="2" face="Comic Sans MS">Categorias 
        cadastradas</font></div></td>
  </tr>

<%  String[] items = categ.montaLista();
	for (int i=0; i<items.length; i++) {
		if (i % 2 == 0) {
%>
	<tr bgcolor="#ffffff"> 
<%  	} else {
%> 
	<tr bgcolor="#C2E9E7"> 
<%		}
%>
		<td width="250"><font face="Comic Sans MS" size="2"><%= items[i] %></font></td>
	</tr>
<%	}
%>
 
</table>

<%@ include file ="inc/base.inc" %> 

</body>
</html>
