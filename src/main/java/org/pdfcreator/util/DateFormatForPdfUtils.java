package org.pdfcreator.util;

import java.text.SimpleDateFormat;
import java.util.Date;

public class DateFormatForPdfUtils {
    private static final SimpleDateFormat SDF = new SimpleDateFormat("yyyyMMdd");

    public static String getCurrentyyyyMMdd(){
        return SDF.format(new Date());
    }
}
