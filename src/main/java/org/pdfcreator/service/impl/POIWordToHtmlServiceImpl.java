package org.pdfcreator.service.impl;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.pdfcreator.context.FileConstants;
import org.pdfcreator.service.POIWordToHtmlService;
import org.pdfcreator.util.PdfFilesUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
@Service
public class POIWordToHtmlServiceImpl implements POIWordToHtmlService {
    private static final Logger logger = LoggerFactory.getLogger(POIWordToHtmlServiceImpl.class);
    private static final String DOC = "doc";
    private static final String DOCX = "docx";
    @Override
    public void wordToHtml(String sourcePath, String targetPath, String picturesPath) {
        String ext = PdfFilesUtils.getFileSuffix(sourcePath);
        String content = null;
        try {
            if (DOC.equals(ext)) {
                HWPFDocument wordDocument = new HWPFDocument(new FileInputStream(sourcePath));
                WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                        DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
                wordToHtmlConverter.setPicturesManager(new PicturesManager() {
                    public String savePicture(byte[] content, PictureType pictureType, String suggestedName,
                                              float widthInches, float heightInches) {
                        String picturesForDocPath = picturesPath + "\\" + ext + "\\" +PdfFilesUtils.getFileNameWithOutSuffix(sourcePath) + "\\" + suggestedName;
                        PdfFilesUtils.createParentDir(picturesForDocPath);
                        FileOutputStream fos = null;
                        try {
                            fos = new FileOutputStream(picturesForDocPath);
                            fos.write(content);
                            fos.close();
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                        return picturesForDocPath;
                    }
                });
                wordToHtmlConverter.processDocument(wordDocument);
                Document htmlDocument = wordToHtmlConverter.getDocument();
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                DOMSource domSource = new DOMSource(htmlDocument);
                StreamResult streamResult = new StreamResult(out);

                TransformerFactory tf = TransformerFactory.newInstance();
                Transformer serializer = tf.newTransformer();
                serializer.setOutputProperty(OutputKeys.ENCODING, FileConstants.DEFAULT_ENCODING);
                serializer.setOutputProperty(OutputKeys.INDENT, "yes");
                serializer.setOutputProperty(OutputKeys.METHOD, "html");
                serializer.transform(domSource, streamResult);
                out.close();
                wordDocument.close();
                PdfFilesUtils.writeFile(new String(out.toByteArray()), targetPath);
                content = out.toString();
                logger.info("doc转html 转换结束...");
            } else if (DOCX.equals(ext)) {
                // 1) 加载word文档生成 XWPFDocument对象
                InputStream in = new FileInputStream(new File(sourcePath));
                XWPFDocument document = new XWPFDocument(in);
                // 2) 解析 XHTML配置 (这里设置IURIResolver来设置图片存放的目录)
                XHTMLOptions options = XHTMLOptions.create();
                String picturesForDocxPath = picturesPath + "\\" + ext + "\\" + PdfFilesUtils.getFileNameWithOutSuffix(sourcePath);
                options.setExtractor(new FileImageExtractor(PdfFilesUtils.createDir(picturesForDocxPath)));
                options.URIResolver(new BasicURIResolver(picturesForDocxPath));
                // 3) 将 XWPFDocument转换成XHTML
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                XHTMLConverter.getInstance().convert(document, baos, options);
                baos.close();
                in.close();
                document.close();
                content = baos.toString();
                PdfFilesUtils.writeFile(content, targetPath);
                logger.info("docx转html 转换结束...");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        content.replace("<html>",FileConstants.HTML_CONTENT_TYPE);
    }

    @Override
    public void wordToHtml(String sourcePath, String targetPath) {

        this.wordToHtml(sourcePath, targetPath, targetPath.substring(0, targetPath.lastIndexOf("\\")));
    }
}
