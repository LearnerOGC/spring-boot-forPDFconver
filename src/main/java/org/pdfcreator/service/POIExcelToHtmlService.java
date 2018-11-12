package org.pdfcreator.service;

public interface POIExcelToHtmlService {
    void excelToHtml(String sourcePath, String targetPath , boolean isWithStyle);

    void excelToHtml(String sourcePath, String targetPath);

}
