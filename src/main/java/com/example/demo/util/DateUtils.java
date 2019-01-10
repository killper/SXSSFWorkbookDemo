package com.example.demo.util;

import java.text.SimpleDateFormat;
import java.util.Date;


public class DateUtils {

  private DateUtils() {

  }

  /*
     *@param: [date, format]
     *@return java.lang.String
     *@author lucasliang
     *@date 12/12/2018
     *@Description format date
    */
  public static String formatDate(Date date, String format) {
    SimpleDateFormat sdf = new SimpleDateFormat(format);
    return sdf.format(date);
  }


}
