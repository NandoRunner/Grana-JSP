package grana;

import javax.servlet.http.*;
import java.util.*;
//import java.io.*;

public class Categoria extends GDB{

  private String SQL;
  private Util util = new Util();

  public Categoria() {
  }


// Existe categoria
  public boolean existe(HttpServletRequest req) {

    if (!(util.eValido(req.getParameter("nomeCategoria")))) {
      return false;
    }

    SQL = "SELECT nomeCategoria FROM Categoria " +
          "WHERE nomeCategoria = '" + req.getParameter("nomeCategoria") + "'";

    return existeRegistro(SQL);
  }


// Insere Categoria
  public boolean insere(HttpServletRequest req) {

    if (!(util.eValido(req.getParameter("nomeCategoria")))) {
      return false;
    }

    SQL = "INSERT INTO Categoria (nomeCategoria) VALUES('" +
          req.getParameter("nomeCategoria") + "')";

    if (!insereRegistro(SQL)) {
      System.out.println("Erro no cadastro da categoria!!!");
      return false;
    } else {
      return true;
    }
  }


// Monta lista de Catagorias
  public String[] montaLista() {

    SQL = "SELECT nomeCategoria FROM Categoria " +
          "WHERE nomeCategoria is not null " +
          "AND not nomeCategoria = '' " +
          "ORDER BY nomeCategoria";

    return geraLista(SQL, 1);
  }


// Monta tabela para relatorio de categorias
  public Vector montaRelatorio() {

    SQL = "SELECT C.nomeCategoria, format(Sum(G.valor), '#,###.00') AS Total  " +
          "FROM gasto AS G, despesa AS D, categoria AS C " +
          "WHERE G.idDespesa = D.idDespesa AND D.idCategoria = C.idCategoria " +
          "GROUP BY C.nomeCategoria " +
          "ORDER BY C.nomeCategoria";

    return montaVector(SQL);
  }


// Monta vector
  public Vector montaVector(String SQL) {

    Vector v = new Vector();

    try {
      rs = stm.executeQuery(SQL);
      while (rs.next()) {
        TabCategoria tab = new TabCategoria();
        tab.setNome(rs.getString(1));
        tab.setValor(rs.getString(2));
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
