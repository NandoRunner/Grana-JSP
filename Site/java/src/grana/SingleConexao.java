package grana;

import java.sql.*;

public class SingleConexao {

  private static Connection con;

  private SingleConexao() {
    try{
      Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
      con = DriverManager.getConnection("jdbc:odbc:grana");
    }
    catch (ClassNotFoundException e)
    {
      System.out.println(e.getMessage());
    }
    catch (SQLException e)
    {
      System.out.println(e.getMessage());
    }
  }

  public static Connection getCon() {
    return con;
  }

  private static final SingleConexao single = new SingleConexao();

  static public SingleConexao instance() {
    return single;
  }
}