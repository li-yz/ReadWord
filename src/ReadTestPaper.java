/**
 * Created by Liyanzhen on 2017/3/6.
 */

import com.common.util.DateFormatUtil;
import com.common.util.FileUploadPathConfig;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.*;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 *

 * @Description:Word�Ծ��ĵ�ģ�ͻ�����

 * @author <a href="mailto:thoslbt@163.com">Thos</a> 42  * @ClassName: WordToHtml 44  * @version V1.0
 *
 */
public class ReadTestPaper {

    /**
     * �س���ASCII��
     */
    private static final short ENTER_ASCII = 13;

    /**
     * �ո��ASCII��
     */
    private static final short SPACE_ASCII = 32;

    /**
     * ˮƽ�Ʊ��ASCII��
     */
    private static final short TABULATION_ASCII = 9;

    public static String htmlText = "";
    public static String htmlTextTbl = "";
    public static int counter=0;
    public static int beginPosi=0;
    public static int endPosi=0;
    public static int beginArray[];
    public static int endArray[];
    public static String htmlTextArray[];
    public static boolean tblExist=false;

    public static final String inputFile="F:\\���߿���ϵͳ\\test2.doc";
    public static final String htmlFile="F:/abc.html";

    public static void main(String argv[])
    {
        try {
            getWordAndStyle(inputFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * ��ȡÿ��������ʽ
     *
     * @param fileName
     * @throws Exception
     */


    public static void getWordAndStyle(String fileName) throws Exception {
        FileInputStream in = new FileInputStream(new File(fileName));
        HWPFDocument doc = new HWPFDocument(in);

        int num=100;

        beginArray=new int[num];
        endArray=new int[num];
        htmlTextArray=new String[num];

        // ȡ���ĵ����ַ�������
        int length = doc.characterLength();

        //���ÿһ�е�����
        List<String> rows = new ArrayList<>();

        // ������ʱ�ַ���,�ü����ж�һ���ַ��Ƿ������ͬ��ʽ

        int cur=0;

        String tempString = "";
        for (int i = 0; i < length - 1; i++) {
            // ��ƪ���µ��ַ�ͨ��һ�����ַ������ж�,rangeΪ�õ��ĵ��ķ�Χ
            Range range = new Range(i, i + 1, doc);

            CharacterRun cr = range.getCharacterRun(0);


            Range range2 = new Range(i + 1, i + 2, doc);
            // �ڶ����ַ�
            CharacterRun cr2 = range2.getCharacterRun(0);
            char c = cr.text().charAt(0);

            // �ж��Ƿ�Ϊ�ո��
            if (c == SPACE_ASCII)
                tempString += " ";
                // �ж��Ƿ�Ϊˮƽ�Ʊ��
            else if (c == TABULATION_ASCII)
                tempString += "\t";
            // �Ƚ�ǰ��2���ַ��Ƿ������ͬ�ĸ�ʽ
            boolean flag = compareCharStyle(cr, cr2);
            if (flag&&c !=ENTER_ASCII)
                tempString += cr.text();
            else {
                //��c�ǻ��з������س�����ʱ��֤��һ�ж�ȡ��ϣ�tempString�б���ļ���һ������
                if(tempString !="" && !tempString.equals("END")){
                    rows.add(tempString);
                }
                tempString = "";
            }

        }

        //word�Ծ�����ģ�ͻ�
        analysisHtmlString(htmlText);
        System.out.println("------------WordToHtmlģ�ͻ��ɹ�----------------");
        analysisRowOfPaper(rows);
    }

    public static void analysisRowOfPaper(List<String>rows){
        //��һ�����Ծ�ı���
        String title = rows.get(0);

        int singleNum =0;//��ѡ������
        int singleScore=0;//��ѡ��ÿ��ķ���
        int multipleNum=0;//��ѡ������
        int multipleScore=0;//��ѡ��ÿ��ķ���
        int judgeNum=0;//�ж�������
        int judgeScore=0;//�ж���ÿ��ķ���
        /***********�Ծ�������ݸ�ֵ*********************/
        for (int i = 0; i < rows.size(); i++) {
            String delHtml = rows.get(i);
            if(delHtml.contains("����ѡ��")){
                String numScore=numScore(delHtml);
                singleNum= Integer.parseInt(numScore.split(",")[0]) ;
                singleScore=Integer.parseInt(numScore.split(",")[1]) ;
            }else if(delHtml.contains("����ѡ��")){
                String numScore=numScore(delHtml);
                multipleNum= Integer.parseInt(numScore.split(",")[0]) ;
                multipleScore=Integer.parseInt(numScore.split(",")[1]) ;
            }else if(delHtml.contains("���ж���")){
                String numScore=numScore(delHtml);
                judgeNum= Integer.parseInt(numScore.split(",")[0]) ;
                judgeScore=Integer.parseInt(numScore.split(",")[1]) ;
            }
        }

        int counter = 0;
        int i=1;
        while(i < rows.size()){
            if(rows.get(i).contains("���ж���") || counter < judgeNum){//��ʼ�����ж��⣬��Ϊ�涨word�ĵ���һ���ж���ռ4��
                while (rows.get(i)==""){
                    i++;
                }
                if(isBigTilete(rows.get(i))){
                    i++;
                }

                String smallTitle = rows.get(i++);
                String option1 = rows.get(i++);
                String option2 = rows.get(i++);
                String rightAnser = rows.get(i++);
                counter++;
            }else if (rows.get(i).contains("����ѡ��") || (counter-judgeNum) < singleNum){
                while (rows.get(i)==""){
                    i++;
                }
                if(isBigTilete(rows.get(i))){
                    i++;
                }

                String smallTitle = rows.get(i++);
                String optionA = rows.get(i++);
                String optionB = rows.get(i++);
                String optionC =  rows.get(i++);
                String optionD = rows.get(i++);
                String rightAnswer =  rows.get(i++);
                counter++;
            }else if(rows.get(i).contains("����ѡ��") || (counter-judgeNum-singleNum) < multipleNum){
                while (rows.get(i)==""){
                    i++;
                }
                if(isBigTilete(rows.get(i))){
                    i++;
                }

                String smallTitle = rows.get(i++);
                String optionA = rows.get(i++);
                String optionB = rows.get(i++);
                String optionC =  rows.get(i++);
                String optionD = rows.get(i++);
                String rightAnswer =  rows.get(i++);
                counter++;
            }else {
                i++;
            }
        }

    }

    public static boolean compareCharStyle(CharacterRun cr1, CharacterRun cr2)
    {
        boolean flag = false;
        if (cr1.isBold() == cr2.isBold() && cr1.isItalic() == cr2.isItalic() && cr1.getFontName().equals(cr2.getFontName())
                && cr1.getFontSize() == cr2.getFontSize()&& cr1.getColor() == cr2.getColor())
        {
            flag = true;
        }
        return flag;
    }

    /*** ������ɫģ��start ********/
    public static int red(int c) {
        return c & 0XFF;
    }

    public static int green(int c) {
        return (c >> 8) & 0XFF;
    }

    public static int blue(int c) {
        return (c >> 16) & 0XFF;
    }

    public static int rgb(int c) {
        return (red(c) << 16) | (green(c) << 8) | blue(c);
    }

    public static String rgbToSix(String rgb) {
        int length = 6 - rgb.length();
        String str = "";
        while (length > 0) {
            str += "0";
            length--;
        }
        return str + rgb;
    }


    public static String getHexColor(int color) {
        color = color == -1 ? 0 : color;
        int rgb = rgb(color);
        return "#" + rgbToSix(Integer.toHexString(rgb));
    }
    /** ������ɫģ��end ******/

    /**
     * д�ļ�
     *
     * @param s
     */
    public static void writeFile(String s) {
        FileOutputStream fos = null;
        BufferedWriter bw = null;
        PrintWriter writer = null;
        try {
            File file = new File(htmlFile);
            fos = new FileOutputStream(file);
            bw = new BufferedWriter(new OutputStreamWriter(fos));
            bw.write(s);
            bw.close();
            fos.close();
            //����ת��
            writer = new PrintWriter(file, "GB2312");
            writer.write(s);
            writer.flush();
            writer.close();
        } catch (FileNotFoundException fnfe) {
            fnfe.printStackTrace();
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }

    }

    /**
     * ����html
     * @param s
     */
    public static void analysisHtmlString(String s){

        String q[] = s.split("<br/>");

        LinkedList<String> list = new LinkedList<String>();

        //������ַ�
        for (int i = 0; i < q.length; i++) {
            if(StringUtils.isNotBlank(q[i].toString().replaceAll("</?[^>]+>","").trim())){

                list.add(q[i].toString().trim());
            }
        }
        String[] result = {};
        String ws[]=list.toArray(result);
        int singleScore = 0;
        int multipleScore = 0;
        int fillingScore = 0;
        int judgeScore = 0;
        int askScore = 0;
        int singleNum = 0;
        int multipleNum = 0;
        int fillingNum = 0;
        int judgeNum = 0;
        int askNum = 0;
        /***********�Ծ�������ݸ�ֵ*********************/
        for (int i = 0; i < ws.length; i++) {
            String delHtml=ws[i].toString().replaceAll("</?[^>]+>","").trim();//ȥ��html
            if(delHtml.contains("����ѡ��")){
                String numScore=numScore(delHtml);
                singleNum= Integer.parseInt(numScore.split(",")[0]) ;
                singleScore=Integer.parseInt(numScore.split(",")[1]) ;
            }else if(delHtml.contains("��������")){
                String numScore=numScore(delHtml);
                multipleNum= Integer.parseInt(numScore.split(",")[0]) ;
                multipleScore=Integer.parseInt(numScore.split(",")[1]) ;
            }else if(delHtml.contains("�������")){
                String numScore=numScore(delHtml);
                fillingNum= Integer.parseInt(numScore.split(",")[0]) ;
                fillingScore=Integer.parseInt(numScore.split(",")[1]) ;
            }else if(delHtml.contains("���ж���")){
                String numScore=numScore(delHtml);
                judgeNum= Integer.parseInt(numScore.split(",")[0]) ;
                judgeScore=Integer.parseInt(numScore.split(",")[1]) ;
            }else if(delHtml.contains("���ʴ���")){
                String numScore=numScore(delHtml);
                askNum= Integer.parseInt(numScore.split(",")[0]) ;
                askScore=Integer.parseInt(numScore.split(",")[1]) ;
            }

        }
        /**************word�Ծ�����ģ�ͻ�****************/
        List<Map<String, Object>> bigTiMaps = new ArrayList<Map<String,Object>>();
        List<Map<String, Object>> smalMaps = new ArrayList<Map<String,Object>>();
        List<Map<String, Object>> sleMaps = new ArrayList<Map<String,Object>>();
        String htmlText="";
        int smalScore=0;
        for (int j = ws.length-1; j>=0; j--) {
            String html= ws[j].toString().trim();//html��ʽ
            String delHtml=ws[j].toString().replaceAll("</?[^>]+>","").trim();//ȥ��html
            if(!isSelecteTitele(delHtml)&&!isTitele(delHtml)&&!isBigTilete(delHtml)){//��
                if(isTitele(delHtml)){
                    smalScore=itemNum(delHtml);
                }
                htmlText=html+htmlText;
            }else if(isSelecteTitele(delHtml)){//ѡ����ѡ����
                Map<String, Object> sleMap = new HashMap<String, Object>();//ѡ����ѡ����
                sleMap.put("seleteItem", delHtml.substring(0, 1));
                sleMap.put("seleteQuest", html+htmlText);
                sleMaps.add(sleMap);
            }else if(isTitele(delHtml)){//С����
                Map<String, Object> smalMap = new HashMap<String, Object>();//С����
                smalMap.put("smalTilete", html+htmlText);
                smalMap.put("smalScore", smalScore>0?smalScore+"":itemNum(delHtml)+"");
                smalMap.put("sleMaps", sleMaps);
                smalMaps.add(smalMap);
            }else if(isBigTilete(delHtml)){//�����
                Map<String, Object> bigTiMap = new HashMap<String, Object>();//�����
                bigTiMap.put("bigTilete", delHtml.substring(2, 5));
                bigTiMap.put("smalMaps", smalMaps);
                bigTiMaps.add(bigTiMap);
            }

        }
        //System.out.println(bigTiMaps.toString());
    }

    //��ȡ����-��Ŀ�����Լ���Ŀ�ܼƷ���
    public static String numScore(String delHtml){

        String regEx="[^0-9+��|,+^0-9]";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(delHtml);
        String s=m.replaceAll("").trim();
        if(StringUtils.isNotBlank(s)){
            if(s.contains(",")){
                return s;
            }else if(s.contains("��")){
                return s.replace("��", ",");
            }else{
                return "0,0";
            }
        }else{
            return "0,0";
        }

    }
    //��ȡÿС�����
    public static int itemNum(String delHtml){
        Pattern pattern = Pattern.compile("��(.*?)��"); //��������
        Matcher matcher = pattern.matcher(delHtml);
        if (matcher.find()&&isNumeric(matcher.group(1))){
            return Integer.parseInt(matcher.group(1));
        }else {
            return 0;
        }
    }
    //�ж�Str�Ƿ��� ����
    public static boolean isNumeric(String str){
        Pattern pattern = Pattern.compile("[0-9]*");
        return pattern.matcher(str).matches();
    }
    //�ж�Str�Ƿ����С�����
    public static boolean isTitele(String str){
        Pattern pattern = Pattern.compile("^([\\d]+[-\\��].*)");
        return pattern.matcher(str).matches();
    }
    //�ж�Str�Ƿ���ѡ����ѡ����
    public static boolean isSelecteTitele(String str){
        Pattern pattern = Pattern.compile("^([a-zA-Z]+[-\\��].*)");
        return pattern.matcher(str).matches();
    }
    //�ж�Str�Ƿ��Ǵ����
    public static boolean isBigTilete(String str){
        boolean iso= false ;
        if(str.contains("һ��")){
            iso=true;
        }else if(str.contains("����")){
            iso=true;
        }else if(str.contains("����")){
            iso=true;
        }else if(str.contains("�ġ�")){
            iso=true;
        }else if(str.contains("�塢")){
            iso=true;
        }else if(str.contains("����")){
            iso=true;
        }else if(str.contains("�ߡ�")){
            iso=true;
        }else if(str.contains("�ˡ�")){
            iso=true;
        }
        return iso;
    }
}
