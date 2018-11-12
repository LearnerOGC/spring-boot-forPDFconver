package org.pdfcreator.service;

public interface WkHtmlToPdfService {
    void htmlConvertToPdf(String sourcePath, String targetPath);

    /**
     * 源文件非html文件转换
     * @param initialFileType 源文件的文件类型
     * @param sourcePath
     * @param targetPath
     */
    void htmlConvertToPdf(String initialFileType, String sourcePath, String targetPath);
}
