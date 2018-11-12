import com.itextpdf.text.pdf.BaseFont;
import org.xhtmlrenderer.pdf.ITextFontResolver;
import org.xhtmlrenderer.pdf.ITextRenderer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

public class Test1 {
    public static boolean convertHtmlToPdf(String sourcePath, String targetPath)
            throws Exception {
        //OutputStream os = new FileOutputStream(outputFile);
        File outPutFile = new File(targetPath);
        File parent = outPutFile.getParentFile();
        //如果pdf保存路径不存在，则创建路径
        if(!parent.exists()){
            parent.mkdirs();
        }

        OutputStream os = new FileOutputStream(targetPath);
        ITextRenderer renderer = new ITextRenderer();
//        String url = new File(sourcePath).toURI().toURL().toString();
//        System.out.println("url:"+url);
//        renderer.setDocument(url);

        // 解决中文支持问题
        ITextFontResolver fontResolver = renderer.getFontResolver();
        fontResolver.addFont("C:/Windows/Fonts/SIMSUN.TTC", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        //解决图片的相对路径问题
        renderer.getSharedContext().setBaseURL("file:C:\\D_files\\poi\\");
        renderer.layout();
        renderer.createPDF(os);

        os.flush();
        os.close();
        return true;
    }
    public static void main(String[] args) {
        try{
            convertHtmlToPdf("C:/D_files/poi/html/docx001.html", "C:/D_files/iText/pdf/docx001.pdf");
        }catch (Exception e){
            e.printStackTrace();
        }

    }
}
