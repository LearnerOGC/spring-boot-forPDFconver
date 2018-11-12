package org.pdfcreator.service;

public interface PdfFileCreatorService {
    void pdfCreator(String sourcePath, String storagePath);

    void pdfCreator(String sourcePath, String tempPath, String storagePath);

    void htmlToPdf(String sourcePath, String storagePath);

    void officeToPdf(String sourcePath, String tempPath, String storagePath);

    void officeToPdf(String sourcePath, String storagePath);

    void txtToPdf(String sourcePath, String storagePath);

}
