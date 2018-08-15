<script language="JavaScript">

function validaCampo(campo, titulo) {
 if (campo.value == "") {
    campo.focus();
	//campo.select();
	alert(titulo + " em branco!");
  	return false;
  } else {
  	return true;
  }
}

</script>