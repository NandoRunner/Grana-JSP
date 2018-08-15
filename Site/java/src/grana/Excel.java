package grana;

import java.util.*;

public class Excel extends GDB{

  private String SQL;
  private Util util = new Util();

  public Excel() {
  }


  public Vector montaRelatorio(String mes) {

    SQL = "SELECT D.nomeDespesa, Sum(G.valor) as ValTotal, count(*) as Qtd, " +
          "ano, ";
	if (!(mes.equals("13"))) {
		SQL = SQL + "mes, ";	 
	}
	SQL = SQL + " FORMAT(diaPadrao, '00') as dd, receita " +
          "FROM gasto AS G, despesa AS D " +
          "WHERE G.idDespesa=D.idDespesa " +
          "and D.especial = true ";
	if (!(mes.equals("13"))) {
		SQL = SQL + "AND mes = " + mes;	 
	}
	SQL = SQL + " GROUP BY D.nomeDespesa, FORMAT(diaPadrao, '00'), ano, ";
	if (!mes.equals("13")) {
		SQL = SQL + "mes, ";	 
	}
	SQL = SQL + " receita " +
    " UNION SELECT C.nomeCategoria, Sum(G.valor) as ValTotal, count(*) as Qtd, " +
          "ano, ";
	if (!(mes.equals("13"))) {
		SQL = SQL + "mes, ";	 
	}
	SQL = SQL + "'99' as dd, false " +
          "FROM gasto AS G, despesa AS D, categoria AS C " +
          "WHERE G.idDespesa=D.idDespesa " +
          "AND D.idCategoria=C.idCategoria " +
          "AND D.especial = false ";
	if (!(mes.equals("13"))) {
		SQL = SQL + "AND mes = " + mes;	 
	}
	SQL = SQL + " GROUP BY  C.nomeCategoria, ano ";
	if (!(mes.equals("13"))) {
		SQL = SQL + ", mes ";	 
	}
	SQL = SQL + " ORDER BY dd";

    Vector v = new Vector();

    try {
      rs = stm.executeQuery(SQL);

      while (rs.next()) {
        TabExcel tab = new TabExcel();
        tab.setNome(rs.getString("nomeDespesa"));
        tab.setValor(rs.getFloat("valTotal"));
        tab.setQtd(rs.getString("Qtd"));
        tab.setDia(rs.getString("dd"));
		if (!(mes.equals("13"))) {
			tab.setMes(rs.getString("mes"));
		}
        tab.setReceita(rs.getBoolean("receita"));
		v.addElement(tab);
      }

      rs.close();
      return v;

    } catch (Exception e) {
      System.out.println(e.getMessage());
      return v;
    }

  }

}
