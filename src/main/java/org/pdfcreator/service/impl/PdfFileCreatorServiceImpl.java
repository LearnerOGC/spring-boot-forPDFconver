package org.pdfcreator.service.impl;

import org.apache.commons.lang3.StringUtils;
import org.pdfcreator.context.FileConstants;
import org.pdfcreator.service.*;
import org.pdfcreator.util.DateFormatForPdfUtils;
import org.pdfcreator.util.EnvironmentUtils;
import org.pdfcreator.util.PdfFilesUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;

@Service
public class PdfFileCreatorServiceImpl implements PdfFileCreatorService {
    private static final Logger logger = LoggerFactory.getLogger(PdfFileCreatorServiceImpl.class);
    @Autowired
    private EnvironmentUtils environmentUtils;
    @Autowired
    private WkHtmlToPdfService wkHtmlToPdfService;
    @Autowired
    private POIWordToHtmlService poiWordToHtmlService;
    @Autowired
    private POIExcelToHtmlService poiExcelToHtmlService;
    @Autowired
    private POIPptToHtmlService poiPptToHtmlService;
    @Autowired
    private TxtToPdfService txtToPdfService;

    private static final String PDF = "pdf";

    @Override
    public void pdfCreator(String sourcePath, String storagePath) {
        this.pdfCreator(sourcePath, sourcePath+"\\temp",storagePath);
    }

    @Override
    public void pdfCreator(String sourcePath, String tempPath, String storagePath) {
        File souFile = new File(sourcePath);
        if(!souFile.exists()){
            logger.error("文件源路径不存在");
            throw new RuntimeException("文件源路径不存在");
        }
        String currentDate = DateFormatForPdfUtils.getCurrentyyyyMMdd();
        String targetPath = storagePath+"\\"+currentDate;
        tempPath = tempPath+"\\"+currentDate;
        Map<String,String> tm = fileListSort(sourcePath, targetPath, FileConstants.PDF_CONVERT_FORMATS);
        Set<String> set = tm.keySet();
        Iterator<String> it = set.iterator();
        //可以通过目标文件夹获取最新的pdf文件是否在源文件夹中存在同名，若存在，则可以从该文件时间点后的文件开始转换
        //遍历treeMap
        logger.info("**********开始遍历文件集合**********");
        while (it.hasNext()) {
            String souPath = it.next();
            String fileType = tm.get(souPath);
            if(StringUtils.isBlank(souPath)){
                logger.error("文件的源路径或文件类型为空");
            }
            logger.info("当前文件路径:"+souPath);
            if(FileConstants.FILES_TYPE_OF_OFFICE.indexOf(fileType) !=  -1){
                String tempFilePath =  tempPath + "\\" + PdfFilesUtils.getFileNameWithOutSuffix(souPath) + FileConstants.FILES_TYPE_OF_TEMP;
                String absolutePath = null;
                //先通过office的相关service转换为html文件
                if(FileConstants.WORD_FILE_TYPE.indexOf(fileType) != -1){
                    poiWordToHtmlService.wordToHtml(souPath, tempFilePath);
                }else if(FileConstants.EXCEL_FILE_TYPE.indexOf(fileType) != -1){
                    poiExcelToHtmlService.excelToHtml(souPath, tempFilePath);
                }else if(FileConstants.PPT_FILE_TYPE.indexOf(fileType) != -1){
                    poiPptToHtmlService.pptToHtml(souPath, tempFilePath);
                }else{
                    logger.error("非office格式文件 不支持转换!");
                }
                //再通过wkhtmltopdf将html转换为pdf文件
                wkHtmlToPdfService.htmlConvertToPdf(fileType, tempFilePath, targetPath);
            }else if(FileConstants.FILES_TYPE_OF_HTML.indexOf(fileType) != -1){
                wkHtmlToPdfService.htmlConvertToPdf(souPath, targetPath);
            }else if(FileConstants.TXT_FILE_TYPE.indexOf(fileType) != -1){
                txtToPdfService.txtToPdf(souPath, targetPath);
            }else{
                logger.error("文件不支持pdf类型转换");
            }

        }

        logger.info("需要转换的文件已全部转换完毕");
    }

    public void htmlToPdf(String sourcePath, String storagePath) {
        File souFile = new File(sourcePath);
        if(!souFile.exists()){
            logger.error("文件源路径不存在");
            throw new RuntimeException("文件源路径不存在");
        }
        String currentDate = DateFormatForPdfUtils.getCurrentyyyyMMdd();
        String targetPath = storagePath+"\\"+currentDate;
        Map<String,String> tm = fileListSort(sourcePath, targetPath, FileConstants.FILES_TYPE_OF_HTML);
        Set<String> set = tm.keySet();
        Iterator<String> it = set.iterator();
        //可以通过目标文件夹获取最新的pdf文件是否在源文件夹中存在同名，若存在，则可以从该文件时间点后的文件开始转换
        //遍历treeMap
        logger.info("**********开始遍历文件集合**********");
        while (it.hasNext()) {
            String souPath = it.next();
            if(StringUtils.isBlank(souPath)){
                logger.error("文件的源路径或文件类型为空");
            }
            logger.info("当前文件路径:"+souPath);
            wkHtmlToPdfService.htmlConvertToPdf(souPath, targetPath);
        }
        logger.info("需要转换的文件已全部转换完毕");
    }

    @Override
    public void officeToPdf(String sourcePath, String storagePath) {
        this.officeToPdf(sourcePath, sourcePath+"\\"+ FileConstants.RELAY_FOLDER_NAME, storagePath);
    }

    public void officeToPdf(String sourcePath, String tempPath, String storagePath) {
        File souFile = new File(sourcePath);
        if(!souFile.exists()){
            logger.error("文件源路径不存在");
            throw new RuntimeException("文件源路径不存在");
        }
        String currentDate = DateFormatForPdfUtils.getCurrentyyyyMMdd();
        String targetPath = storagePath+"\\"+currentDate;
        tempPath = tempPath+"\\"+currentDate;
        Map<String,String> tm = fileListSort(sourcePath, targetPath, FileConstants.FILES_TYPE_OF_OFFICE);
        Set<String> set = tm.keySet();
        Iterator<String> it = set.iterator();
        //可以通过目标文件夹获取最新的pdf文件是否在源文件夹中存在同名，若存在，则可以从该文件时间点后的文件开始转换
        //遍历treeMap
        logger.info("**********开始遍历文件集合**********");
        while (it.hasNext()) {
            String souPath = it.next();
            String tempFilePath =  tempPath + "\\" + PdfFilesUtils.getFileNameWithOutSuffix(souPath) + FileConstants.FILES_TYPE_OF_TEMP;
            String fileType = tm.get(souPath);
            if(StringUtils.isBlank(souPath)){
                logger.error("文件的源路径或文件类型为空");
            }
            String absolutePath = null;
            System.out.println("当前文件路径:"+souPath);
            //先通过office的相关service转换为html文件
            if(FileConstants.WORD_FILE_TYPE.indexOf(fileType) != -1){
                poiWordToHtmlService.wordToHtml(souPath, tempFilePath);
            }else if(FileConstants.EXCEL_FILE_TYPE.indexOf(fileType) != -1){
                poiExcelToHtmlService.excelToHtml(souPath, tempFilePath);
            }else if(FileConstants.PPT_FILE_TYPE.indexOf(fileType) != -1){
                poiPptToHtmlService.pptToHtml(souPath, tempFilePath);
            }else{
                logger.error("非office格式文件 不支持转换!");
            }
            //再通过wkhtmltopdf将html转换为pdf文件
            wkHtmlToPdfService.htmlConvertToPdf(fileType, tempFilePath, targetPath);
        }
        logger.info("需要转换的文件已全部转换完毕");
    }

    @Override
    public void txtToPdf(String sourcePath, String storagePath) {
        File souFile = new File(sourcePath);
        if(!souFile.exists()){
            logger.error("文件源路径不存在");
            throw new RuntimeException("文件源路径不存在");
        }
        String currentDate = DateFormatForPdfUtils.getCurrentyyyyMMdd();
        String targetPath = storagePath+"\\"+currentDate;
        Map<String,String> tm = fileListSort(sourcePath, targetPath, FileConstants.TXT_FILE_TYPE);
        Set<String> set = tm.keySet();
        Iterator<String> it = set.iterator();
        //可以通过目标文件夹获取最新的pdf文件是否在源文件夹中存在同名，若存在，则可以从该文件时间点后的文件开始转换
        //遍历treeMap
        logger.info("**********开始遍历文件集合**********");
        while (it.hasNext()) {
            String souPath = it.next();
            if(StringUtils.isBlank(souPath)){
                logger.error("文件的源路径或文件类型为空");
            }
            logger.info("当前文件路径:"+souPath);
            txtToPdfService.txtToPdf(souPath, targetPath);
        }
        logger.info("需要转换的文件已全部转换完毕");
    }

    /**
     * 根据源路径获取对应文件类型最近修改时间文件的绝对路径与对应的文件类型
     * @return
     */
    private Map<String,String> fileListSort(String sourcePath, String targetPath, String filesType) {
        Map<String,String> souTm = new HashMap<>();
        File soufile = new File(sourcePath);
        File sourcesFile[] = soufile.listFiles();
        if(sourcesFile == null || sourcesFile.length <= 0){
            logger.error("源文件不存在");
        }
        int souFileNum = sourcesFile.length;
        //拿到所有的文件类型数组
        String[] targetPaths = filesType.split(",");
        Map<String, String> fileMap = new HashMap<>();
        Map<String, Long> lastModifyTimeMap = new HashMap<>();
        //循环数组 最终生成以 Key-文件类型,Value-该文件类型文件夹下源文件的最后修改时间 的Map
        for(int i=0; i<targetPaths.length; i++){
            TreeMap<Long,String> tarTm = new TreeMap<Long,String>();
            //找到对应目标文件夹
            File tarFile = new File(targetPath+"\\"+targetPaths[i]+"topdf");
            if(tarFile.exists()){
                //目标文件夹存在，则获取目标文件夹下所有文件名(不包含文件后缀)并生成TreeMap(Key-时间的Long类型 Value-文件名)
                File targetFiles[] = tarFile.listFiles();
                if(targetFiles != null && targetFiles.length > 0) {
                    int tarFileNum = targetFiles.length;
                    for (int j = 0; j < tarFileNum; j++) {
                        if(targetFiles[j].exists() && !targetFiles[j].isDirectory() && PDF.equals(PdfFilesUtils.getFileSuffix(targetFiles[j].getAbsolutePath()))){
                            Long tempLong = new Long(targetFiles[j].lastModified());
                            tarTm.put(tempLong, PdfFilesUtils.getFileNameWithOutSuffix(targetFiles[j].getAbsolutePath()));
                        }else{
                            logger.error("目标文件不存在或文件类型与目标文件夹不符");
                        }

                    }
                }
            }
            //若TreeMap中有值(即目标文件夹下已存在转换过的文件)则获取TreeMap中最近更新过的文件名-tarTm.lastKey()
            //并根据该文件名加上源文件的路径生成对应的源文件，若存在，则拿到该源文件的最近更新时间并存入Map中(Key-文件类型 Value-Long类型的时间)
            if(tarTm != null && tarTm.size() > 0) {
                //将目标文件夹中最近更新的文件名存入map中
                fileMap.put(targetPaths[i], tarTm.get(tarTm.lastKey()));
                Set<String> set = fileMap.keySet();
                Iterator<String> it = set.iterator();
                while (it.hasNext()) {
                    Object key = it.next();
                    String objValue = fileMap.get(key);
                    File soursFile = new File(sourcePath +"\\"+ objValue + "." + key);
                    if (soursFile.exists() && !soursFile.isDirectory()) {
                        lastModifyTimeMap.put(key.toString(), soursFile.lastModified());
                    } else {
                        lastModifyTimeMap.put(key.toString(), 0L);
                    }
                }
            }
        }
        //循环所有源文件，若文件类型符合支持转换类型，则根据上述所生成的Map获取对应文件类型的最近更新时间
        //并与当前源文件的最近更新时间比较，大于等于Map中保存的时间，则需要转换，并存入到Map中(Key-文件的绝对路径 Value-文件的后缀)
        for (int i = 0; i < souFileNum; i++) {
            Long tempLong = new Long(sourcesFile[i].lastModified());
            if(sourcesFile[i].exists() && sourcesFile[i].getName().indexOf(".") != -1 && !sourcesFile[i].isDirectory()) {
                //获取文件类型
                String fileType = PdfFilesUtils.getFileSuffix(sourcesFile[i].getAbsolutePath());
                //过滤掉不支持转换的文件类型 且仅当文件的最后修改时间大于且等于记录的源文件中最近更新同名文件的时间
                if(filesType.indexOf(fileType) != -1 && sourcesFile[i].lastModified() >= (lastModifyTimeMap.get(fileType) == null ? 0L : lastModifyTimeMap.get(fileType))){
                    souTm.put(sourcesFile[i].getAbsolutePath(), fileType);
                    logger.info("--需要转换的源文件的绝对路径:"+sourcesFile[i].getAbsolutePath()+"文件名:"+sourcesFile[i].getName() +"文件最后修改时间:"+sourcesFile[i].lastModified());
                }
            }

        }
        return souTm;
    }

}
