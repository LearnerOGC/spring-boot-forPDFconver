import org.apache.poi.hslf.usermodel.*;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;


/**
 * ppt转html
 * @author Red_Ant
 * 20180925
 */
public class PptToHtml {

    private static void pptToPicture(String tempContextUrl, String outPath){
        //文件夹名
        List<String> imgList = new ArrayList<String>();
        File file = new File(tempContextUrl);
        File folder = new File(outPath + File.separator + "20180925");
        try {
            folder.mkdirs();
            FileInputStream is = new FileInputStream(file);
            HSLFSlideShow ppt = new HSLFSlideShow(is);
            is.close();
            Dimension pgsize = ppt.getPageSize();
            List<HSLFSlide> slide = ppt.getSlides();
            for (int i = 0; i < slide.size(); i++) {
                for (HSLFShape shape : ppt.getSlides().get(i).getShapes()) {
                    if (shape instanceof HSLFTextShape) {
                        HSLFTextShape tsh = (HSLFTextShape) shape;
                        for (HSLFTextParagraph p : tsh) {
                            for (HSLFTextRun r : p) {
                                r.setFontIndex(1);
                                r.setFontFamily("宋体");
                            }
                        }
                    }
                }
                BufferedImage img = new BufferedImage(pgsize.width,pgsize.height, BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();
                graphics.setPaint(Color.BLUE);
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
                slide.get(i).draw(graphics);
                String imgName = File.separator + "20180925" + File.separator +"pict_"+ (i + 1) + ".jpeg";
                FileOutputStream out = new FileOutputStream(outPath + imgName);
                javax.imageio.ImageIO.write(img, "jpeg", out);
                out.close();
                imgList.add("20180925" + File.separator +"pict_"+ (i + 1) + ".jpeg");
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        createPPTHtml(outPath,imgList, tempContextUrl);
    }

    /**自己创建的html代码
     * 原理上就是，把上一步ppt转的图片
     * 以html的方式呈现出来
     */
    public static void createPPTHtml(String wordPath,List<String> imgList,String tempContextUrl){
        StringBuilder sb = new StringBuilder("<!doctype html><html><head><meta charset='utf-8'><title>无标题文档</title></head><body><div align=\"center\">");
        if (imgList != null && !imgList.isEmpty()) {
            for (String img : imgList) {
                sb.append("<img src='" + img + "' /><br>");
            }
        }
        sb.append("</div></body></html>");
        try {
            File file = new File(wordPath + File.separator + "20180925.html");
            BufferedWriter bufferedWriter = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(file),"UTF-8"));
            bufferedWriter.write(sb.toString());
            bufferedWriter.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        pptToPicture("C:\\D_files\\make_files\\word\\T03 Spring AOP.ppt",
                "C:\\D_files\\make_files\\word\\20181111");
    }

}