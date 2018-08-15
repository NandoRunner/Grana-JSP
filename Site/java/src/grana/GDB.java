package grana;

import java.sql.*;
import java.util.*;


public class GDB {

  private static Connection con;
  protected PreparedStatement pstm;
  protected Statement stm;
  protected ResultSet rs;


  public GDB() {
    SingleConexao single = SingleConexao.instance();
    con = single.getCon();

    try{
      stm = con.createStatement();
    } catch (SQLException se) {
      se.printStackTrace();
    }
  }


// Insere registro
  public boolean insereRegistro(String SQL){

    try{
      stm.executeUpdate(SQL);
      return true;

    } catch (Exception e) {
      System.out.println(e.getMessage());
      return false;
    }
  }

// gera lista
  public String[] geraLista(String SQL, int campos) {

    Vector v = new Vector();
    String[] s2 = new String[0];

    try {
      rs = stm.executeQuery(SQL);

      while (rs.next()) {
        for(int i = 1; i <= campos; i++)
          v.addElement(rs.getString(i));
      }
      String[] s = new String[v.size()];
      v.copyInto(s);

      rs.close();
      return s;

    } catch (Exception e) {
      System.out.println(e.getMessage());
      return s2;
    }
  }

// busca ID
  public int buscaID(String SQL, String chave) {

    try {
      rs = stm.executeQuery(SQL);

      if (rs.next()) {
        return rs.getInt(chave);
      } else {
        return 0;
      }

    } catch (Exception e) {
      System.out.println(e.getMessage());
      return 0;
    }

  }

// existe registro
  public boolean existeRegistro(String SQL) {

    try {
      rs = stm.executeQuery(SQL);

      if (rs.next()) {
        return true;
      } else {
        return false;
      }

    } catch (Exception e) {
      System.out.println(e.getMessage());
      return false;
    }

  }


}

