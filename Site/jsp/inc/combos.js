<script language="JavaScript">

function geraComboMes(mes, c) {
	var m = new Array(13);
	m[1] = 'Janeiro';
	m[2] = 'Fevereiro';
	m[3] = 'Março';
	m[4] = 'Abril';
	m[5] = 'Maio';
	m[6] = 'Junho';
	m[7] = 'Julho';
	m[8] = 'Agosto';
	m[9] = 'Setembro';
	m[10] = 'Outubro';
	m[11] = 'Novembro';
	m[12] = 'Dezembro';
	m[13] = '(Todos)';
	document.write('<select name="' + mes + '" onChange="javascript:testaMes()">');
	for(i=1; i<=13; i++) {
		if(c != i)
			document.write('<option value="' + i + '">' + m[i]);
		else
			document.write('<option value="' + i + '" selected>' + m[i]);
	}
	document.write('</select>');
}

</script>