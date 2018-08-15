package grana;

public class TabCategoria {

  private String nome;
  private String valor;
  private String mes;
  private String ano;

  public TabCategoria() {
  }

  public void setNome(String s) {
    nome = s;
  }

  public void setValor(String s) {
    valor = s;
  }

  public void setMes(String s) {
    mes = s;
  }

  public void setAno(String s) {
    ano = s;
  }

  public String getNome() {
    return nome;
  }

  public String getValor() {
    return valor;
  }

  public String getMes() {
    return mes;
  }

  public String getAno() {
    return ano;
  }

}