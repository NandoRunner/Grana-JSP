package grana;

public class TabGasto {

  private String nomeDespesa;
  private float valor;
  private String dia;
  private boolean receita;
  private boolean especial;

  public TabGasto() {
  }

  public void setNomeDespesa(String s) {
    nomeDespesa = s;
  }

  public void setValor(float f) {
    valor = f;
  }

  public void setDia(String s) {
    dia = s;
  }

  public void setReceita(boolean b) {
    receita = b;
  }

  public void setEspecial(boolean b) {
    especial = b;
  }

  public String getNomeDespesa() {
    return nomeDespesa;
  }

  public float getValor() {
    return valor;
  }

  public String getDia() {
    return dia;
  }

  public boolean getReceita() {
    return receita;
  }

  public boolean getEspecial() {
    return especial;
  }

}
