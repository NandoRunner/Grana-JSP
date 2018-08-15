package grana;

import javax.servlet.http.*;
import java.util.*;

public class Gasto extends GDB{

  private String SQL;
  private Util util = new Util();

  public Gasto() {
  }


// Insere
  public boolean insere(HttpServletRequest req) {

    if (!(util.eValido(req.getParameter("val"))))
      return false;

    SQL = "SELECT idDespesa FROM Despesa WHERE nomeDespesa ='" +
      req.getParameter("idDespesa") + "'";

	int id = buscaID(SQL, "idDespesa");

    if (id == 0)
      return false;

    SQL = "INSERT INTO Gasto (idDespesa, valor, dia, mes, ano, info) VALUES(" +
       id + ", " +
       util.valorF((String)req.getParameter("val")) + ", " +
       req.getParameter("dia") + ", " +
       req.getParameter("mes") + ", " +
       req.getParameter("ano") + ", '" +
       req.getParameter("info") + "')";

    
	return insereRegistro(SQL);
  }

  public Vector montaLista(String mes) {

    SQL = "SELECT nomeDespesa, valor, dia, receita, especial " +
          "FROM Gasto,  Despesa " +
          "WHERE Gasto.idDespesa = Despesa.idDespesa " +
		  "AND mes = " + mes +
          " ORDER BY mes desc, dia desc, nomeDespesa";

    Vector v = new Vector();

    try {
      rs = stm.executeQuery(SQL);
      while (rs.next()) {
        TabGasto tab = new TabGasto();
        tab.setNomeDespesa(rs.getString("nomeDespesa"));
        tab.setValor(rs.getFloat("valor"));
        tab.setDia(rs.getString("dia"));
        tab.setReceita(rs.getBoolean("receita"));
        tab.setEspecial(rs.getBoolean("especial"));
        v.addElement(tab);
      }

      rs.close();
      return v;

    } catch (Exception e) {
      System.out.println(e.getMessage());
      return v;
    }

  }

  public String totalLista(String mes) {

    SQL = "SELECT format(sum(valor), '#,###.00') as val, (dia " +
          " & '/' & mes) as dt, receita, especial, valor " +
          "FROM Gasto,  Despesa " +
          "WHERE Gasto.idDespesa = Despesa.idDespesa " +
          "AND mes = " + mes +
          " ORDER BY mes desc, dia desc, nomeDespesa";

	String str = "";

    try {
      rs = stm.executeQuery(SQL);
      if (rs.next()) {
		str = rs.getString("val");
	  } else {
		str = "";
	  }
		rs.close();
	  return str;


    } catch (Exception e) {
      System.out.println(e.getMessage());
      return "";
    }

  }


}
