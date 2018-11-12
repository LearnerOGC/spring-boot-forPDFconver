package org.pdfcreator.service;

public interface POIPptToHtmlService {
    void pptToHtml(String sourcePath, String targetPath);

    void pptToHtml(String sourcePath, String targetPath, String imagePath);
}
