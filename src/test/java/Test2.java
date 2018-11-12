import com.itextpdf.text.pdf.BaseFont;
import org.xhtmlrenderer.pdf.ITextFontResolver;
import org.xhtmlrenderer.pdf.ITextRenderer;

import java.io.*;

public class Test2 {
    public static boolean convertHtmlToPdf(String sourcePath, String targetPath)
            throws Exception {
        //OutputStream os = new FileOutputStream(outputFile);
        File outPutFile = new File(targetPath);
        File parent = outPutFile.getParentFile();
        //如果pdf保存路径不存在，则创建路径
        if(!parent.exists()){
            parent.mkdirs();
        }

        //读取html
        FileInputStream fis =new FileInputStream(sourcePath);
        StringWriter writers = new StringWriter();
        InputStreamReader isr = null;
        String string = null;
        //此处将io流转换成String
        try {
            isr = new InputStreamReader(fis,"utf-8");//包装基础输入流且指定编码方式
            //将输入流写入输出流
            char[] buffer = new char[2048];
            int n = 0;
            while (-1 != (n = isr.read(buffer))) {
                writers.write(buffer, 0, n);
            }
        }catch (Exception e){
            e.printStackTrace();
        } finally {
            if (isr != null)
                try {
                    isr.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
        }
        if (writers!=null){
            string = writers.toString();
        }
        string = string.replace("&nbsp;","&#160;");
        System.out.print(string);

        OutputStream os = new FileOutputStream(targetPath);
        ITextRenderer renderer = new ITextRenderer();
//        String url = new File(sourcePath).toURI().toURL().toString();
//        System.out.println("url:"+url);
//        renderer.setDocument(url);
        renderer.setDocumentFromString(string);
        // 解决中文支持问题
        ITextFontResolver fontResolver = renderer.getFontResolver();
        fontResolver.addFont("C:/Windows/Fonts/SIMSUN.TTC", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        //解决图片的相对路径问题
        renderer.getSharedContext().setBaseURL("file:C:\\D_files\\make_files\\word\\temp\\20181111\\");
        renderer.layout();
        renderer.createPDF(os);

        os.flush();
        os.close();
        return true;
    }
    public static void main(String[] args) {
        try{
            convertHtmlToPdf("C:\\D_files\\make_files\\word\\temp\\20181111\\doc001.html", "C:\\D_files\\make_files\\word\\temp\\20181111\\doc001.pdf");
        }catch (Exception e){
            e.printStackTrace();
        }

    }
}
