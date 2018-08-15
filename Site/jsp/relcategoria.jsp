<%@ page import="java.util.*" %>
<%@ page import="grana.*" %>
<html>
<head>
<title>Grana 2003 for Web - Relatório de Categorias (consolidado)</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<link href="grana.css" rel="stylesheet" type="text/css">
</head>

<jsp:useBean id="categ" scope="session" class="grana.Categoria" />
<jsp:setProperty name="categ" property="*" />

<body>
<table width="550" height="20" border="0" align="center" cellpadding="10" cellspacing="0">
  <tr> 
    <td height="22" align="center" valign="middle" bgcolor="#009999"><font color="#FFFFFF" size="4" face="Comic Sans MS">Relat&oacute;rio 
      de categorias </font></td>
  </tr>
</table>
<br>
<form name="frm" action="categoria.jsp" method="post" onSubmit="return testaCampos()">
  <br>
</form>
<table width="330" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr bgcolor="#009999"> 
    <td width="250"><div align="left"><font color="#FFFFFF" size="2" face="Comic Sans MS">Categoria</font></div></td>
    <td width="80"><div align="right"><font color="#FFFFFF" size="2" face="Comic Sans MS">Valor</font></div></td>
  </tr>
  <%  Vector tab =  categ.montaRelatorio();
  	for (int i=0; i<tab.size(); i++) {
		if (i % 2 == 0) {
%>
  <tr bgcolor="#ffffff"> 
    <%  	} else {
%>
  <tr bgcolor="#C2E9E7"> 
    <%		}
%>
    <td width="250"><font face="Comic Sans MS" size="2"> <%=((TabCategoria)tab.elementAt(i)).getNome() %></td>
    <td width="80"><div align="right"><font face="Comic Sans MS" size="2"> <%=((TabCategoria)tab.elementAt(i)).getValor() %> 
      </div></td>
  </tr>
  <% }
%>
</table>

<%@ include file ="inc/base.inc" %>

</body>
</html>
