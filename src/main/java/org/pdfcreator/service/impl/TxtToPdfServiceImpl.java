package org.pdfcreator.service.impl;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.commons.lang3.StringUtils;
import org.pdfcreator.context.FileConstants;
import org.pdfcreator.service.TxtToPdfService;
import org.pdfcreator.util.PdfFilesUtils;
import org.springframework.stereotype.Service;

import java.io.*;
@Service
public class TxtToPdfServiceImpl implements TxtToPdfService {
    private static final String TXT = "txt";
    @Override
    public void txtToPdf(String sourcePath, String targetPath) {
        String fileType = PdfFilesUtils.getFileSuffix(sourcePath);
        if(StringUtils.isBlank(fileType) || !TXT.equals(fileType)){
            return ;
        }
        targetPath = targetPath + "\\" + fileType + "topdf\\" + PdfFilesUtils.getFileNameWithOutSuffix(sourcePath) + ".pdf";
        Document document = new Document();
        try {
            PdfFilesUtils.createParentDir(targetPath);
            OutputStream os = new FileOutputStream(new File(targetPath));
            PdfWriter.getInstance(document, os);
            document.open();
            BaseFont baseFont = BaseFont.createFont(FileConstants.TXT_FONT_PATH, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            Font font = new Font(baseFont);
            InputStreamReader isReader = new InputStreamReader(new FileInputStream(new File(sourcePath)), FileConstants.DEFAULT_ENCODING);
            BufferedReader reader = new BufferedReader(isReader);
            String line = null;
            while ((line = reader.readLine()) != null){
                document.add(new Paragraph(line, font));
            }
            document.close();
            isReader.close();
            os.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }catch (DocumentException de){
            de.printStackTrace();
        }catch (IOException ioe){
            ioe.printStackTrace();
        }

    }
}
