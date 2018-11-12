package org.pdfcreator.service.impl;

import org.apache.poi.hslf.usermodel.*;
import org.apache.poi.xslf.usermodel.*;
import org.pdfcreator.context.FileConstants;
import org.pdfcreator.service.POIPptToHtmlService;
import org.pdfcreator.util.PdfFilesUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

@Service
public class POIPptToHtmlServiceImpl implements POIPptToHtmlService {
    private static final Logger logger = LoggerFactory.getLogger(POIPptToHtmlServiceImpl.class);
    private static final String PPT = "ppt";
    private static final String PPTX = "pptx";

    @Override
    public void pptToHtml(String sourcePath, String targetPath, String imagePath) {
        File pptFile = new File(sourcePath);
        if (pptFile.exists()) {
            try {
                String type = PdfFilesUtils.getFileSuffix(sourcePath);
                String pptFileName = pptFile.getName().substring(0, pptFile.getName().lastIndexOf("."));

                if (PPT.equals(type)) {
                    String htmlStr = toImage2003(sourcePath, targetPath, pptFileName, imagePath);
                    PdfFilesUtils.writeFile(htmlStr, targetPath);
                } else if (PPTX.equals(type)) {
                    String htmlStr = toImage2007(sourcePath, targetPath, pptFileName, imagePath);
                    PdfFilesUtils.writeFile(htmlStr, targetPath);
                } else {
                    logger.info("不是ppt文件");
                }

            } catch (Exception e) {
                e.printStackTrace();
            }
        } else {
            logger.info("文件不存在");
        }
    }

    @Override
    public void pptToHtml(String sourcePath, String targetPath) {
        this.pptToHtml(sourcePath, targetPath, targetPath.substring(0, targetPath.lastIndexOf("\\")));
    }

    private String toImage2007(String sourcePath, String targetPath, String pptFileName, String pptImagePath) throws Exception {
        String htmlStr = "";
        FileInputStream is = new FileInputStream(sourcePath);
        XMLSlideShow ppt = new XMLSlideShow(is);
        PdfFilesUtils.createParentDir(targetPath);// create html dir
        Dimension pgsize = ppt.getPageSize();
        System.out.println(pgsize.width + "--" + pgsize.height);
        StringBuffer sb = new StringBuffer();
        sb.append(FileConstants.HTML_CONTENT_TYPE);
        sb.append("</head>");
        for (int i = 0; i < ppt.getSlides().size(); i++) {
            try {
                // 防止中文乱码
                for (XSLFShape shape : ppt.getSlides().get(i).getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape tsh = (XSLFTextShape) shape;
                        for (XSLFTextParagraph p : tsh) {
                            for (XSLFTextRun r : p) {
                                r.setFontFamily(FileConstants.PPT_TO_HTML_FONT_FAMILY);
                            }
                        }
                    }
                }
                BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();
                // clear the drawing area
                graphics.setPaint(Color.white);
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
                // render
                ppt.getSlides().get(i).draw(graphics);
                // save the output
                String imageDir = pptImagePath + "\\" + "ppt\\" + pptFileName + "\\";
                PdfFilesUtils.createDir(imageDir);// create image dir
                String imagePath = imageDir + pptFileName + "-" + (i + 1) + "." + FileConstants.PPT_IMAGE_TYPE;
                sb.append("<br>");
                sb.append("<img src=" + "\"" + imagePath + "\"" + "/>");
                FileOutputStream out = new FileOutputStream(imagePath);
                javax.imageio.ImageIO.write(img, FileConstants.PPT_IMAGE_TYPE, out);
                out.close();
            } catch (Exception e) {
                System.out.println("第" + i + "张ppt转换出错");
            }
        }
        is.close();
        ppt.close();
        System.out.println("success");
        sb.append("</html>");
        htmlStr = sb.toString();
        return htmlStr;
    }

    private String toImage2003(String sourcePath, String targetPath, String pptFileName, String pptImagePath) {
        String htmlStr = "";
        HSLFSlideShow ppt = null;
        try {
            ppt = new HSLFSlideShow(new HSLFSlideShowImpl(sourcePath));
            PdfFilesUtils.createParentDir(targetPath);// create html dir
            Dimension pgsize = ppt.getPageSize();
            StringBuffer sb = new StringBuffer();
            sb.append(FileConstants.HTML_CONTENT_TYPE);
            sb.append("</head>");
            for (int i = 0; i < ppt.getSlides().size(); i++) {
                // 防止中文乱码
                for (HSLFShape shape : ppt.getSlides().get(i).getShapes()) {
                    if (shape instanceof HSLFTextShape) {
                        HSLFTextShape tsh = (HSLFTextShape) shape;
                        for (HSLFTextParagraph p : tsh) {
                            for (HSLFTextRun r : p) {
                                r.setFontFamily(FileConstants.PPT_TO_HTML_FONT_FAMILY);
                            }
                        }
                    }
                }
                BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();
                // clear the drawing area
                graphics.setPaint(Color.white);
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
                // render
                ppt.getSlides().get(i).draw(graphics);
                String imageDir = pptImagePath + "\\" + "ppt\\" + pptFileName + "\\";
                PdfFilesUtils.createDir(imageDir);// create image dir
                String imagePath = imageDir + pptFileName + "-" + (i + 1) + ".jpeg";
                sb.append("<br>");
                sb.append("<img src=" + "\"" + imagePath + "\"" + "/>");
                FileOutputStream out = new FileOutputStream(imagePath);
                javax.imageio.ImageIO.write(img, FileConstants.PPT_IMAGE_TYPE, out);
                out.close();
            }
            ppt.close();
            System.out.println("success");
            sb.append("</html>");
            htmlStr = sb.toString();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return htmlStr;
    }

}
