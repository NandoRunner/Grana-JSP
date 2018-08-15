<html>
<head>
<title>Grana 2003 for Web - Relat&oacute;rio Excel</title>
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
<%@ include file ="inc/combos.js" %>

<script language="JavaScript">

function testaMes() {
	
	document.frm.fMes.value = document.frm.mes.value;
	document.frm.submit();
}

</script>

<link href="grana.css" rel="stylesheet" type="text/css">
</head>

<jsp:useBean id="excel" scope="session" class="grana.Excel" />
<jsp:setProperty name="excel" property="*" />

<jsp:useBean id="util" scope="session" class="grana.Util" />
<jsp:setProperty name="util" property="*" />

<body>
<table width="550" height="20" border="0" align="center" cellpadding="10" cellspacing="0">
  <tr> 
    <td height="22" align="center" valign="middle" bgcolor="#009999"><font color="#FFFFFF" size="4" face="Comic Sans MS">Relat&oacute;rio 
      Excel</font></td>
  </tr>
</table>
<br><div align="center"><br>
<form name="frm" action="relexcel.jsp" method="post">
  <input name="fMes" type="hidden" value="<%=session.getValue("Mes") %>">

<script type="text/javascript">
	geraComboMes("mes", document.frm.fMes.value);
</script>
</div><br>

</form>


<table width="410" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr bgcolor="#009999"> 
    <td width="40"><div align="center"><font color="#FFFFFF" size="2" face="Comic Sans MS">Dia</font></div></td>
    <td width="250"><font color="#FFFFFF" size="2" face="Comic Sans MS">Gasto</font></td>
    <td width="80"><div align="right"><font color="#FFFFFF" size="2" face="Comic Sans MS">Valor</font></div></td>
    <td width="40"><div align="right"><font color="#FFFFFF" size="2" face="Comic Sans MS">Qtd</font></div></td>
  </tr>
  <%  Vector rel = excel.montaRelatorio((String)session.getValue("Mes"));
	TabExcel tab;
	String dia;
	String cor;	     			
    float 	total = 0;
	float	rec=0;
	float	des=0;	
	for (int i=0; i<rel.size(); i++) {
		tab = (TabExcel)rel.elementAt(i);
		dia = tab.getDia();
		
		if (tab.getReceita()) {
			total = total + tab.getValor();
			rec = rec + tab.getValor();
			cor = "#0000CC";
		} else {
			if ((tab.getNome()).equals("Receita")) {
				total = total + tab.getValor();
				rec = rec + tab.getValor();
				cor = "#0000CC";
			} else {
				total = total - tab.getValor();
				des = des + tab.getValor();
				cor = "#000000";
			}
		}
		if (dia.equals("99") || dia.equals("00")) 
			dia = "";
		else
			tab.setQtd("");
			
		if (i % 2 == 0) {
%>
  <tr bgcolor="#ffffff"> 
    <%  	} else {
%>
  <tr bgcolor="#C2E9E7"> 
    <%		}
%>
    <td width="40"><div align="center"><font color="#0000CC" size="1" face="Comic Sans MS"> 
        <%=dia %> </div></td>
    <td width="250"><font face="Comic Sans MS" color="<%=cor %>" size="2"> <%=tab.getNome() %></td>
    <td width="80"><div align="right"><font face="Comic Sans MS" color="<%=cor %>" size="2"> 
        <%=util.valorStr(tab.getValor()) %> </div></td>
    <td width="40"><div align="right"><%=tab.getQtd() %></div></td>
  </tr>
  <% }
%>
  <tr> 
    <td width="40"></td>
    <td width="250"><font face="Comic Sans MS" color="#0000CC" size="2"> Receita</td>
    <td width="80"><div align="right"><font face="Comic Sans MS" color="#0000CC" size="2"> 
        <%=util.valorStr(rec) %></div></td>
    <td width="40">&nbsp;</td>
  </tr>
  <tr> 
    <td width="40"></td>
    <td width="250"><font face="Comic Sans MS" color="#0000CC" size="2"> <font color="#990000">Despesa</font></td>
    <td width="80"><div align="right"><font face="Comic Sans MS" color="#990000" size="2"> 
        <%=util.valorStr(des) %></div></td>
    <td width="40">&nbsp;</td>
  </tr>
  <tr> 
    <td width="40"></td>
    <td width="250"><font face="Comic Sans MS" color="#0000CC" size="2"> TOTAL</td>
    <td width="80"><div align="right"><font face="Comic Sans MS" color="#0000CC" size="2"> 
        <%=util.valorStr(total) %></div></td>
    <td width="40">&nbsp;</td>
  </tr>
</table>

<%@ include file ="inc/base.inc" %>

</body>
</html>
