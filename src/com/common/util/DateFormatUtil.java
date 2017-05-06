package com.common.util;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by Liyanzhen on 2017/5/6.
 */
public class DateFormatUtil {
    public static String formatDate(Date date){
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        String str = sdf.format(date);
        return str;
    }
}
