package org.pdfcreator.util;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.util.regex.Pattern;

@Service
public class EnvironmentUtils {
    public static final String LINUX_SYSTEM_NAME = "linux";
    public static final String WINDOWS_SYSTEM_NAME = "windows";
    @Value("${wkhtmltopdf.path}")
    private String WKHTMLTOPDF_PATH;


    public String getWkhtmlToPdfExePath(){
        return WKHTMLTOPDF_PATH;
    }

    public String getCurrentSystem(){
        String osName = System.getProperty("os.name");
        if (Pattern.matches("Linux.*", osName)) {
            //获取linux系统下WkhtmlToPdf主程序的位置
            return LINUX_SYSTEM_NAME;
        } else if (Pattern.matches("Windows.*", osName)) {
            //获取windows系统下WkhtmlToPdf主程序的位置
            return WINDOWS_SYSTEM_NAME;
        }
        return null;
    }

}
