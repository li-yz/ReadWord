import com.common.util.DateFormatUtil;

import java.util.Date;

/**
 * Created by Liyanzhen on 2017/5/6.
 */
public class Test {
    public static void main(String[] args){
        Date date = new Date();
        System.out.println(DateFormatUtil.formatDate(date));
    }
}
