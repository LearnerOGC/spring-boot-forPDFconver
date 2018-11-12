import org.pdfcreator.context.FileConstants;
import org.pdfcreator.util.PdfFilesUtils;

import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Test3 {
    public static void main(String[] args) {
//        String sou = "C:\\D_files\\make_files\\doc001.doc";
//        System.out.println(sou.substring(sou.lastIndexOf("\\")+1, sou.lastIndexOf(".")) + FileConstants.FILES_TYPE_OF_TEMP);
//        System.out.println(sou.substring(0, sou.lastIndexOf("\\")));
//        System.out.println("pdf".equals("excel001.pdf".substring("excel001.pdf".lastIndexOf(".")+1)));
//        File file = new File("C:\\D_files\\make_files\\temp\\123\\1234.txt");
//        //如果pdf保存路径不存在，则创建路径
//        if(!file.exists()){
//            file.mkdirs();
//
//        }
        String path = "C:\\D_files\\make_files\\word\\doc001234.html";
        File file = new File(path);
//        String path = "C:\\D_files\\make_files\\word\\temp\\20181109\\doc\\doc001\\0.png";
        StringBuilder sb = new StringBuilder();
        try {
            BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(file)));
            String line = null;
            while((line = reader.readLine()) != null){
                sb.append(line);
            }
            reader.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }catch (IOException ioe){
            ioe.printStackTrace();
        }
        String htmlFile = sb.toString();
        System.out.println("转换前:"+htmlFile);
        StringBuffer sbu = new StringBuffer();
        Pattern p = Pattern.compile("font-family:[^><;\"]*;");
        Matcher m = p.matcher(htmlFile);
        while(m.find()){
            String temp = m.group();
            String replaseStr = "font-family:'宋体';";
            m.appendReplacement(sbu, replaseStr);
        }
        m.appendTail(sbu);
        System.out.println("转换后:"+sbu);

        String ss = "{margin-left:0.33333334in;margin-top:0.06944445in;margin-bottom:0.06944445in;text-align:start;hyphenate:auto;font-family:新宋体;font-size:21pt;}";
        String par = "font-family:[^(><;\")]*;";
        System.out.println("是否匹配 :"+Pattern.matches(par,ss));

//        m.replaceAll("font-family:'宋体';");
//        PdfFilesUtils.writeFile(htmlFile,path);
    }

}
