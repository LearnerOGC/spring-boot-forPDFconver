package org.pdfcreator.service;

public interface POIWordToHtmlService {
    void wordToHtml(String sourcePath, String targetPath, String picturesPath);

    void wordToHtml(String sourcePath, String targetPath);

}
