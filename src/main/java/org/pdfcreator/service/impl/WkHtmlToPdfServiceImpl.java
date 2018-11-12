package org.pdfcreator.service.impl;

import org.apache.commons.lang3.StringUtils;
import org.pdfcreator.context.FileConstants;
import org.pdfcreator.service.WkHtmlToPdfService;
import org.pdfcreator.util.EnvironmentUtils;
import org.pdfcreator.util.PdfFilesUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.*;
@Service
public class WkHtmlToPdfServiceImpl implements WkHtmlToPdfService {
    private static final Logger logger = LoggerFactory.getLogger(WkHtmlToPdfServiceImpl.class);

    @Autowired
    private EnvironmentUtils environmentUtils;

    @Override
    public void htmlConvertToPdf(String sourcePath, String targetPath) {
        this.htmlConvertToPdf(null ,sourcePath, targetPath);
    }

    @Override
    public void htmlConvertToPdf(String initialFileType, String sourcePath, String targetPath) {
        String environment = environmentUtils.getCurrentSystem();
        String wkPath = environmentUtils.getWkhtmlToPdfExePath();
        if(StringUtils.isBlank(initialFileType)){
            initialFileType = PdfFilesUtils.getFileSuffix(sourcePath);
        }
        //转换前：先变更html文件的font-family
        PdfFilesUtils.transformFontFamily(sourcePath);
        targetPath = targetPath+"\\"+ initialFileType +"topdf\\" +PdfFilesUtils.getFileNameWithOutSuffix(sourcePath) +".pdf";
        File file = new File(targetPath);
        File parent = file.getParentFile();
        //如果pdf保存路径不存在，则创建路径
        if(!parent.exists()){
            parent.mkdirs();
        }
        StringBuilder cmd = new StringBuilder();
        if(!EnvironmentUtils.WINDOWS_SYSTEM_NAME.equals(environment)){
            //非windows 系统
//            wmUrl = FileUtil.convertSystemFilePath("/home/ubuntu/wkhtmltox/bin/wkhtmltopdf");
//            wmUrl =NtcoPropertyPlaceholderConfigurer.get("wkhtmltopdf_linux");
        }
        cmd.append(wkPath);
        cmd.append(" ");
        cmd.append("--disable-external-links ");//禁止页面中的外链生成超链接
        cmd.append("--disable-internal-links ");//禁止页面中的内链生成超链接
        cmd.append("--load-error-handling ignore ");//指定如何处理无法加载的页面 ignore
        cmd.append("--disable-javascript ");//禁止WEB页面执行 javascript
        cmd.append("--page-size A4 ");//设置纸张大小: A4, Letter, etc.
//        cmd.append(" --header-line ");//页眉下面的线
        //cmd.append(" --header-center 这里是页眉这里是页眉这里是页眉这里是页眉 ");//页眉中间内容
//        cmd.append(" --margin-top 1cm ");//设置页面上边距 (default 10mm)
        // cmd.append(" --header-html file:///"+WebUtil.getServletContext().getRealPath("")+ FileUtil.convertSystemFilePath("\\style\\pdf\\head.html"));// (添加一个HTML页眉,后面是网址)
//        cmd.append(" --header-spacing 1 ");// (设置页眉和内容的距离,默认0)
        //cmd.append(" --footer-center (设置在中心位置的页脚内容)");//设置在中心位置的页脚内容
        // cmd.append(" --footer-html file:///"+WebUtil.getServletContext().getRealPath("")+FileUtil.convertSystemFilePath("\\style\\pdf\\foter.html"));// (添加一个HTML页脚,后面是网址)
//        cmd.append(" --footer-line ");//* 显示一条线在页脚内容上)
//        cmd.append(" --footer-spacing 1 ");// (设置页脚和内容的距离)
        cmd.append(sourcePath);
        cmd.append(" ");
        cmd.append(targetPath);
        Process proc = null;
        try{
            logger.info("开始转换");
            proc = Runtime.getRuntime().exec(cmd.toString());
            HtmlToPdfInterceptor error = new HtmlToPdfInterceptor(proc.getErrorStream());
            HtmlToPdfInterceptor output = new HtmlToPdfInterceptor(proc.getInputStream());
            error.start();
            output.start();
            proc.waitFor();
        }catch(Exception e){
            e.printStackTrace();
        }finally {
            if(proc != null){
                proc.destroy();
                logger.info("转换完成：销毁命令执行器");
            }
        }
    }

    /**
     * Html转换PDF需要读取input与error流数据，避免阻塞
     */
    private class HtmlToPdfInterceptor extends Thread  {
        private InputStream is;

        public HtmlToPdfInterceptor(InputStream is){
            this.is = is;
        }

        public void run(){
            BufferedReader br = null;
            try{
                InputStreamReader isr = new InputStreamReader(is, FileConstants.DEFAULT_ENCODING);
                br = new BufferedReader(isr);
                String line = null;
                while ((line = br.readLine()) != null) {
                    System.out.println(line);
                    br.readLine();
                }
            }catch (IOException e){
                e.printStackTrace();
            }finally {
                if(br != null){
                    try {
                        br.close();
                    }catch (IOException e){
                        e.printStackTrace();
                    }
                }
            }
        }
    }
}
