package grana;

import javax.servlet.http.*;
import java.util.*;


public class Despesa extends GDB{

  private String SQL;
  private Util util = new Util();

  public Despesa() {
  }

  public boolean existe(HttpServletRequest req) {

    if (!(util.eValido(req.getParameter("nomeDespesa")))) {
      return false;
    }

    SQL = "SELECT nomeDespesa FROM Despesa " +
          "WHERE nomeDespesa = '" + req.getParameter("nomeDespesa") + "'";

    if (existeRegistro(SQL)) {
      return true;
    } else {
      return false;
    }

  }


// Insere Despesa
  public boolean insere(HttpServletRequest req) {

    if (!(util.eValido(req.getParameter("nomeDespesa"))))
      return false;

    SQL = "SELECT idCategoria FROM Categoria WHERE nomeCategoria ='" +
          req.getParameter("idCateg") + "'";

    int id = buscaID(SQL, "idCategoria");

    if (id == 0)
      return false;

    boolean receita = ((String)req.getParameter("idCateg")).equals("Receita");


    String valorPadrao = (String)req.getParameter("valorPadrao");
    if (!(util.eValido(valorPadrao)))
      valorPadrao = "0";

    String diaPadrao = (String)req.getParameter("diaPadrao");
    if (!(util.eValido(diaPadrao)))
      diaPadrao = "0";

    SQL = "INSERT INTO Despesa (idCategoria, nomeDespesa, valorPadrao, " +
          "diaPadrao, especial, receita) VALUES(" +
          id + ", '" +
          req.getParameter("nomeDespesa") + "', " +
          util.valorF((String)valorPadrao) + ", " +
          diaPadrao + ", " +
          req.getParameter("especial") + ", " +
          receita + ")";


    if (!insereRegistro(SQL)) {
      System.out.println("Erro no cadastro da Despesa!!!");
      return false;
    } else {
      return true;
    }
  }


  public Vector montaLista() {

    SQL = "SELECT nomeDespesa, valorPadrao FROM Despesa " +
          "WHERE nomeDespesa is not null " +
          "AND not nomeDespesa = '' " +
          "ORDER BY nomeDespesa";


    Vector v = new Vector();

    try {
      rs = stm.executeQuery(SQL);
      while (rs.next()) {
        TabDespesa tab = new TabDespesa();
        tab.setNome(rs.getString(1));
        tab.setValorPadrao(rs.getString(2));
        v.addElement(tab);
      }

      rs.close();
      return v;

    } catch (Exception e) {
      System.out.println(e.getMessage());
      return v;
    }

  }

// Monta tabela para relatorio de categorias
  public Vector montaRelatorio() {

    SQL = "SELECT nomeCategoria, nomeDespesa, valorPadrao, diaPadrao " +
          "FROM Despesa, Categoria " +
          "WHERE Despesa.idCategoria = Categoria.idCategoria " +
          "AND nomeDespesa is not null " +
          "AND not nomeDespesa = '' " +
          "ORDER BY nomeCategoria, nomeDespesa";

    return montaVector(SQL);
  }


// Monta vector
  public Vector montaVector(String SQL) {

    Vector v = new Vector();

    try {
      rs = stm.executeQuery(SQL);
      while (rs.next()) {
        TabDespesa tab = new TabDespesa();
        tab.setNome(rs.getString(2));
        tab.setNomeCategoria(rs.getString(1));
        tab.setValorPadrao(rs.getString(3));
        tab.setDia(rs.getString(4));
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