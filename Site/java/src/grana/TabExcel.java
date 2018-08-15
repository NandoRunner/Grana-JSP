package grana;


public class TabExcel {
  private String	nome;
  private float     valor;
  private String	qtd;
  private String	dia;
  private String	mes;
  private boolean	receita;

  public TabExcel() {
  }

  public void setNome(String s) {
    nome = s;
  }


  public void setValor(float f) {
    valor = f;
  }

  public void setQtd(String s) {
    qtd = s;
  }

  public void setDia(String s) {
    dia = s;
  }

  public void setMes(String s) {
    mes = s;
  }

  public void setReceita(boolean b) {
    receita = b;
  }

  public String getNome() {
    return nome;
  }

  public float getValor() {
    return valor;
  }

  public String getQtd() {
    return qtd;
  }

  public String getDia() {
    return dia;
  }

  public String getMes() {
    return mes;
  }

  public boolean getReceita() {
    return receita;
  }


}