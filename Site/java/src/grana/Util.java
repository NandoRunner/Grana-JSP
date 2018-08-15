package grana;

import java.text.*;

/**
 * <p>Title: </p>
 * <p>Description: </p>
 * <p>Copyright: Copyright (c) 2003</p>
 * <p>Company: </p>
 * @author unascribed
 * @version 1.0
 */

public class Util {

  public Util() {
  }

  public boolean eValido(String s) {

    if (s == null) {
      return false;
    } else if (s.equals("")) {
      return false;
    } else {
      return true;
    }
  }

  public String valorF(String v) {
   v.replace(',', '.');
   return v;
  }

  public String valorStr(float v) {

    NumberFormat nf = NumberFormat.getInstance();
    nf.setMinimumFractionDigits(2);
    nf.setGroupingUsed(true);
    return nf.format(v);

  }

}

