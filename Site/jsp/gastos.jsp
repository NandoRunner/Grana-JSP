<html>
<head>
<title>Grana 2003 for Web - Lançamento de Gastos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<%@ page import="java.util.*" %>
<%@ page import="grana.*" %>
<%
	String strMes = new String();
	if (request.getParameter("nMes") != null)
		strMes = request.getParameter("nMes");
	else if (request.getParameter("fMes") != null)
		strMes = request.getParameter("fMes");

	session.putValue("Mes", strMes);
%>		
<%@ include file ="inc/validaCampo.js" %>
<%@ include file ="inc/combos.js" %>

<script language="JavaScript">

function mostraValor() {

//	document.frm.val.value = document.frm.idDespesa.value;
//	alert(document.frm.val.value);	
}

function testaMes() {
	
	document.frm2.fMes.value = document.frm2.mes.value;
	document.frm2.submit();
}

function testaCampos(){
  
  if (!validaCampo(document.frm.idDespesa, "Despesa"))
  	return false;

  if (!validaCampo(document.frm.val, "Valor"))
  	return false;

  if (!validaCampo(document.frm.dia, "Dia"))
  	return false;

  if (!validaCampo(document.frm.mes, "Mes"))
  	return false;

  if (!validaCampo(document.frm.ano, "Ano"))
  	return false;
	
  return true;	
}
</script>

<link href="grana.css" rel="stylesheet" type="text/css">
</head>

<jsp:useBean id="gasto" scope="session" class="grana.Gasto" />
<jsp:setProperty name="gasto" property="*" />

<jsp:useBean id="util" scope="session" class="grana.Util" />
<jsp:setProperty name="util" property="*" />

<body>
<table width="550" height="20" border="0" align="center" cellpadding="10" cellspacing="0">
  <tr> 
    <td height="22" align="center" valign="middle" bgcolor="#009999"><font color="#FFFFFF" size="4" face="Comic Sans MS">Lan&ccedil;amento 
      de gastos</font></td>
  </tr>
</table>
<br>
<form name="frm" action="gastos.jsp?nMes=<%= new java.util.Date().getMonth() + 1 %>" method="post" onSubmit="return testaCampos()">
  <table width="500" height="119" border="0" align="center" cellpadding="2" cellspacing="5">
    <tr> 
      <td width="130" height="36" align="right" valign="middle" class="unnamed1"><font color="#009999">Despesa</font></td>
      <td height="36" colspan="3" valign="middle" class="unnamed1"> 
	  <select name="idDespesa" id="select" onChange="javascript:mostraValor()">
          <option value=""></option>
          <jsp:useBean id="despesa" scope="session" class="grana.Despesa" />
          <jsp:setProperty name="despesa" property="*" />
          <% 
	Vector lista = despesa.montaLista();
	TabDespesa td;
	for (int i=0; i<lista.size(); i++) {
		td = (TabDespesa)lista.elementAt(i);
%>
          <option value="<%=td.getNome() %>"><%=td.getNome() %></option>
          <%
	}
%>
        </select></td>
    </tr>
    <tr> 
      <td height="36" align="right" valign="middle" class="unnamed1"><font color="#009999">Valor</font></td>
      <td width="84" height="36" valign="middle" class="unnamed1"> 
	  <input name="val" type="text" id="val" size="9" value="" maxlength="7" ></td>
	  
      <td width="54" valign="middle" class="unnamed1"><div align="right"><font color="#009999">Data</font></div></td>
      <td width="191" valign="middle" class="unnamed1"><input name="dia" type="text" id="dia" size="2" value="<%= new java.util.Date().getDate() %>" maxlength="2">
        / 
        <input name="mes" type="text" id="mes" size="2" value="<%= new java.util.Date().getMonth() + 1 %>" maxlength="2">
        / 
        <input name="ano" type="text" id="ano" size="4" value="<%= new java.util.Date().getYear()+ 1900 %>" maxlength="4"></td>
    </tr>
    <tr> 
      <td height="36" align="right"><font color="#009999">Informa&ccedil;&otilde;es</font></td>
      <td height="36" colspan="3"><textarea name="info" cols="30" rows="3" id="valor"></textarea></td>
    </tr>
  </table>
  <br>
  <table width="220" border="0" align="center" cellspacing="5">
    <tr> 
      <td height="26"><div align="center"> 
          <input name="Submit" type="submit" value="Salvar">
        </div></td>
      <td><div align="center"> 
          <input name="Submit2" type="reset" value="Limpar">
        </div></td>
    </tr>
  </table>
  <br>
  <table width="306" height="36" border="0" align="center" cellspacing="5">
    <tr align="left" valign="middle"> 
      <td width="294" height="36" colspan="2"> 
<%		if (gasto.insere(request)) { 
%>
        <p align="center">Gasto lan&ccedil;ado!</p>
        <%		}
%>
    </tr>
  </table>
  </form>
<form name="frm2" action="gastos.jsp" method="post">
  <div align="center"><font color="#000000" size="2" face="Comic Sans MS">Gastos 
    lan&ccedil;ados em</font> 
    <input name="fMes" type="hidden" value="<%=session.getValue("Mes") %>">
    <script type="text/javascript">
	geraComboMes("mes", document.frm2.fMes.value);
</script>
  </div>
</form>
</div>
<table width="410" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr bgcolor="#009999"> 
    <td width="250"><div align="left"><font color="#FFFFFF" size="2" face="Comic Sans MS">Despesa/receita</font></div></td>
    <td width="80"><div align="right"><font color="#FFFFFF" size="2" face="Comic Sans MS">Valor</font></div></td>
    <td width="40"><div align="right"><font color="#FFFFFF" size="2" face="Comic Sans MS">Dia</font></div></td>
  </tr>
<%  lista = gasto.montaLista((String)session.getValue("Mes"));
	TabGasto tab;
	String cor;	     			
 	float 	total = 0;
	float	rec=0;
	float	des=0;	
	
   	for (int i=0; i<lista.size(); i++) {
		tab = (TabGasto)lista.elementAt(i);
		if (tab.getReceita()) {
			total = total + tab.getValor();
			rec = rec + tab.getValor();
			cor = "#0000CC";
		} else {
			total = total - tab.getValor();
			des = des + tab.getValor();
		 	cor = "#000000";
		}
		
		if (i % 2 == 0) {
%>
  <tr bgcolor="#ffffff"> 
<%  	} else {
%>
  <tr bgcolor="#C2E9E7"> 
<%		}
%>
	<td width="250"><font face="Comic Sans MS" color="<%=cor %>" size="2"> <%=tab.getNomeDespesa() %>
      <div align="left"></div></td>
    <td width="80"><div align="right"><font face="Comic Sans MS" color="<%=cor %>" size="2"> <%=util.valorStr(tab.getValor()) %> 
      </div></td>
    <td width="40"><div align="right"><font face="Comic Sans MS" color="<%=cor %>" size="2"> <%=tab.getDia() %> 
      </div></td>
  </tr>
<% 
	}
%>
 <tr> 
    <td width="250"><font face="Comic Sans MS" color="#0000CC" size="2"> Receita</td>
    <td width="80"><div align="right"><font face="Comic Sans MS" color="#0000CC" size="2"> 
        <%=util.valorStr(rec) %></div></td>
    <td width="40">&nbsp;</td>
  </tr>
  <tr> 
    <td width="250"><font face="Comic Sans MS" color="#0000CC" size="2"> <font color="#990000">Despesa</font></td>
    <td width="80"><div align="right"><font face="Comic Sans MS" color="#990000" size="2"> 
        <%=util.valorStr(des) %></div></td>
    <td width="40">&nbsp;</td>
  </tr>
   <tr>	
	<td width="250"><font face="Comic Sans MS" color="#0000CC" size="2"> TOTAL</td>
    <td width="80"><div align="right"><font face="Comic Sans MS" color="#0000CC" size="2">
	<%=util.valorStr(total) %></div></td>
    <td width="40"></td>
  </tr>
</table>

<%@ include file ="inc/base.inc" %>

</body>
</html>
