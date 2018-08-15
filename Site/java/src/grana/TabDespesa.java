package grana;

public class TabDespesa {

  private String nome;
  private String nomeCategoria;
  private String valorPadrao;
  private String dia;

  public TabDespesa() {
  }

  public void setNome(String s) {
    nome = s;
  }

  public void setNomeCategoria(String s) {
    nomeCategoria = s;
  }

  public void setValorPadrao(String s) {
    valorPadrao = s;
  }

  public void setDia(String s) {
    dia = s;
  }

  public String getNome() {
    return nome;
  }

  public String getNomeCategoria() {
    return nomeCategoria;
  }

  public String getValorPadrao() {
    return valorPadrao;
  }

  public String getDia() {
    return dia;
  }

}