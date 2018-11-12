
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
  * @ClassName: PoiExcelToHtmlUtil
  * @Description: TODO(poi转excel为html)
  * @author liang
  * @date 2016年5月5日 下午3:50:20
  * 2017年5月19日，张鹏：添加隐藏行的过滤，修改excelToHtml入口参数，保证逻辑的正确性
  * 2017年5月26日，张鹏：添加图片处理，当excel中包含图片时，先将图片保存到指定位置，后面在html代码中引用该图片，不同使用环境可修改保存位置，在web环境下，可指定到webapp下
  * 2017年5月27日，张鹏：针对合并单元格后又隐藏了首行（此时下面的行列虽然可以显示正确的数据，但是poi无法读取到对应的value）的特殊类型excel，进行合并代码的重构
  * 2017年5月27日，张鹏：边框全部使用黑色基本样式，解决合并后部分边框为灰色的问题；如果有特殊要求，可在dealStyle方法中解放原来的颜色获取方法
  * 2017年6月5日，张鹏：样式处理方法dealstyle中，修改width属性的获取方式，解决由于合并单元格只计算单独宽度时造成的宽度计算错误的问题，并将原有的宽度转换为了pt单位，原有的px方式存在错误
  * 2017年6月6日，张鹏：添加多sheet的excel转换方式，提供此类调用的一套接口
  */
public class PoiExcelToHtmlUtil {
    private static Map<String, Object> map[];
    private static Map<String, PictureForExcelToHtml> wangPicMap = new HashMap<String,  PictureForExcelToHtml>();

    /**
      * 程序入口方法（读取指定位置的excel，将其转换成html形式的字符串，并保存成同名的html文件在相同的目录下，默认带样式）
      * @param realPath 文件路径
      * @param saveName 文件名称
      * @param suffix 文件后缀
      * @return <table>...</table> 字符串
      */
            public static void excelWriteToHtml(String realPath, final String saveName,final String suffix){
        System.out.println("*****excel转html 正在转换...*****");
        String excelFile=realPath + "\\" + saveName + suffix;
        String htmlFile=realPath + "\\" + saveName + ".html";
        File sourcefile = new File(excelFile);
        try {
            InputStream is = new FileInputStream(sourcefile);
            String content = readExcelToHtml(is,true);
            writeFile(content, htmlFile);
            
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("*****excel转html 转换结束...*****");
    }
    
            public static void excelWriteToHtml(String excelFilePath, String htmlFilePath){
        System.out.println("*****excel转html 正在转换...*****");
        File sourcefile = new File(excelFilePath);
        try {
            InputStream is = new FileInputStream(sourcefile);
            String content = readExcelToHtml(is,true);
            writeFile(content, htmlFilePath);
//            PDFUtils.standardHTML(htmlFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("*****excel转html 转换结束...*****");
    }


    /**
      * 程序入口方法（将指定路径的excel文件读取成字符串）
      * @param filePath 文件的路径
      * @param isWithStyle 是否需要表格样式 包含 字体 颜色 边框 对齐方式
      * @return <table>...</table> 字符串
      */
            public static String readExcelToHtml(String filePath , boolean isWithStyle){
        InputStream is = null;
        String htmlExcel = null;
        try {
            File sourcefile = new File(filePath);
            is = new FileInputStream(sourcefile);
            Workbook wb = WorkbookFactory.create(is);
            htmlExcel = readWorkbook(wb,isWithStyle);
        } catch (Exception e) {
            e.printStackTrace();
        }finally{
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return htmlExcel;
    }
    /**
      * 程序入口方法（将指定路径的excel文件读取成字符串）
      * @param is excel转换成的输入流
      * @param isWithStyle 是否需要表格样式 包含 字体 颜色 边框 对齐方式
      * @return <table>...</table> 字符串
      */
            public static String readExcelToHtml(InputStream is,boolean isWithStyle){
        String htmlExcel = null;
        try {
            Workbook wb = WorkbookFactory.create(is);
            htmlExcel = readWorkbook(wb,isWithStyle);
        } catch (Exception e) {
            e.printStackTrace();
        }finally{
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return htmlExcel;
    }

    /**
      * 根据excel的版本分配不同的读取方法进行处理
      * @param wb
      * @param isWithStyle
      * @return
      */
            private static String readWorkbook(Workbook wb,boolean isWithStyle){
        String htmlExcel = "";
        if (wb instanceof XSSFWorkbook) {
            XSSFWorkbook xWb = (XSSFWorkbook) wb;
            htmlExcel = getExcelInfo(xWb,isWithStyle);
        }else if(wb instanceof HSSFWorkbook){
            HSSFWorkbook hWb = (HSSFWorkbook) wb;
            htmlExcel = getExcelInfo(hWb,isWithStyle);
        }
        return htmlExcel;
    }

    /**
      * 读取excel成string
      * @param wb
      * @param isWithStyle
      * @return
      */
        public static String getExcelInfo(Workbook wb, boolean isWithStyle){
        StringBuffer sb = new StringBuffer();
        Sheet sheet = wb.getSheetAt(0);//获取第一个Sheet的内容
        // map等待存储excel图片
        Map<String, PictureData> sheetIndexPicMap = getSheetPictrues(0, sheet, wb);
        //临时保存位置，正式环境根据部署环境存放其他位置
        try {
            if(sheetIndexPicMap!=null)
                printImg(sheetIndexPicMap);
        } catch (IOException e) {
            e.printStackTrace();
        }

        //读取excel拼装html
        int lastRowNum = sheet.getLastRowNum();
            System.out.println("lastRowNum:"+lastRowNum);
        map = getRowSpanColSpanMap(sheet);
//        sb.append("<table style='border-collapse:collapse;width:100%;'>");
        sb.append("<table style='border-collapse:collapse;width:100%;'>");
        Row row = null;        //兼容
        Cell cell = null;    //兼容

        for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
            System.out.println("FirstRowNum:"+rowNum);
            row = sheet.getRow(rowNum);
            int lastColNum = row.getLastCellNum();
            int rowHeight = row.getHeight();
            if (row == null) {
                sb.append("<tr><td >  </td></tr>");
                continue;
            }else if(row.getZeroHeight()){
                continue;
            }else if(rowHeight==0){
                continue;//针对jxl的隐藏行（此类隐藏行只是把高度设置为0，单getZeroHeight无法识别）
            }
            sb.append("<tr>");


            for (int colNum = 0; colNum < lastColNum; colNum++) {
                if(sheet.isColumnHidden(colNum))continue;
                String imageRowNum = "0_"+rowNum+"_"+colNum;
                String imageHtml = "";
                cell = row.getCell(colNum);
                if ((sheetIndexPicMap!=null && !sheetIndexPicMap.containsKey(imageRowNum) || sheetIndexPicMap==null) && cell == null) {    //特殊情况 空白的单元格会返回null+//判断该单元格是否包含图片，为空时也可能包含图片
                    sb.append("<td>  </td>");
                    continue;
                }
                 PictureForExcelToHtml  PictureForExcelToHtml = null;
                if(sheetIndexPicMap!=null && sheetIndexPicMap.containsKey(imageRowNum)){
                    String imagePath = "C:\\D_files\\poi\\html\\pic\\"+imageRowNum+".jpg";
                    imageHtml = "<img style='z-index:100;' src='"+imagePath+"'>";
                }
                String stringValue = null;
                if(cell != null){
                    stringValue = getCellValue(cell);
                }
                if (map[0].containsKey(rowNum + "," + colNum)) {
                    String pointString = (String)map[0].get(rowNum + "," + colNum);
                    int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
                    int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
                    int rowSpan = bottomeRow - rowNum + 1;
                    int colSpan = bottomeCol - colNum + 1;
                    if(map[2].containsKey(rowNum + "," + colNum)){
                        rowSpan = rowSpan-(Integer)map[2].get(rowNum + "," + colNum);
//                        System.out.println(rowNum + "," + colNum+"减少rows:"+map[2].get(rowNum + "," + colNum)+"，剩余"+rowSpan+";当前值："+stringValue);
                    }
                    sb.append("<td rowspan= '" + rowSpan + "' colspan= '"+ colSpan + "' ");
                    if(map.length>3 && map[3].containsKey(rowNum + "," + colNum)){
                        //此类数据首行被隐藏，value为空，需使用其他方式获取值
                        stringValue = getMergedRegionValue(sheet,rowNum,colNum);
                    }
                } else if (map[1].containsKey(rowNum + "," + colNum)) {
                    map[1].remove(rowNum + "," + colNum);
                    continue;
                }else if(isPicColSpanMerge(rowNum,colNum,sb,imageHtml)){
                    
                }
                else {
                    sb.append("<td >");
                }

                //判断是否需要样式
//                if(isWithStyle){
//                    if(cell !=null ){
//                        dealExcelStyle(wb, sheet, cell, sb);//处理单元格样式
//                    }
//                }
//                sb.append(">");
                
                if(sheetIndexPicMap!=null && sheetIndexPicMap.containsKey(imageRowNum)) sb.append(imageHtml);
                if (stringValue == null || "".equals(stringValue.trim())) {
                    sb.append("   ");
                } else {
                    // 将ascii码为160的空格转换为html下的空格（ ）
                    sb.append(stringValue.replace(String.valueOf((char) 160)," "));
                }
                sb.append("</td>");
            }
            sb.append("</tr>");
        }
        sb.append("</table>");
        return sb.toString();
    }
    
            
            public static boolean isPicColSpanMerge(int rowNum,int colNum,StringBuffer sb ,String imageHtml){
        for (Map.Entry<String,  PictureForExcelToHtml> m :wangPicMap.entrySet())  {  
             PictureForExcelToHtml picture = m.getValue();
            if ((picture.getLeft_up_x() <= rowNum && picture.getRight_down_x() >= rowNum)
                    && (picture.getLeft_up_y() <= colNum && picture.getRight_down_y()>= colNum) ) {
                    if(picture.getLeft_up_x() == rowNum && picture.getLeft_up_y() == colNum){
                        int rowSpan = picture.getRight_down_x() - picture.getLeft_up_x();
                        int colSpan = picture.getRight_down_y() - picture.getLeft_up_y();
                        sb.append("<td rowspan= '" + rowSpan + "' colspan= '"+ colSpan + "' >");
                        return true;
                    }
            }
        }  
        
        return false;
    }
    
        //    public static String getExcelInfo(Workbook wb, boolean isWithStyle){
//        StringBuffer sb = new StringBuffer();
//        Sheet sheet = wb.getSheetAt(0);//获取第一个Sheet的内容
//        // map等待存储excel图片
//        Map<String, PictureData> sheetIndexPicMap = getSheetPictrues(0, sheet, wb);
//        //临时保存位置，正式环境根据部署环境存放其他位置
//        try {
//            if(sheetIndexPicMap!=null)
//                printImg(sheetIndexPicMap);
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//
//        //读取excel拼装html
//        int lastRowNum = sheet.getLastRowNum();
//        map = getRowSpanColSpanMap(sheet);
//        sb.append("<table style='border-collapse:collapse;width:100%;'>");
//        Row row = null;        //兼容
//        Cell cell = null;    //兼容
//
//        for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
//            if(rowNum>1000) break;
//            row = sheet.getRow(rowNum);
//            int lastColNum = row.getLastCellNum();
//            int rowHeight = row.getHeight();
//            if (row == null) {
//                sb.append("<tr><td >  </td></tr>");
//                continue;
//            }else if(row.getZeroHeight()){
//                continue;
//            }else if(rowHeight==0){
//                continue;//针对jxl的隐藏行（此类隐藏行只是把高度设置为0，单getZeroHeight无法识别）
//            }
//            sb.append("<tr>");
//
//
//            for (int colNum = 0; colNum < lastColNum; colNum++) {
//                if(sheet.isColumnHidden(colNum))continue;
//                String imageRowNum = "0_"+rowNum+"_"+colNum;
//                String imageHtml = "";
//                cell = row.getCell(colNum);
//                if ((sheetIndexPicMap!=null && !sheetIndexPicMap.containsKey(imageRowNum) || sheetIndexPicMap==null) && cell == null) {    //特殊情况 空白的单元格会返回null+//判断该单元格是否包含图片，为空时也可能包含图片
//                    sb.append("<td>  </td>");
//                    continue;
//                }
//                if(sheetIndexPicMap!=null && sheetIndexPicMap.containsKey(imageRowNum)){
////                    String imagePath = "D:\\pic"+imageRowNum+".jpeg";
//                    String imagePath = "/home/dev/pics/pic"+imageRowNum+".jpg";
////                    imageHtml = "<img src='"+imagePath+"' style='height:"+rowHeight/20+"px;'>";
//                    imageHtml = "<img style='z-index:100;' src='"+imagePath+"'>";
//                }
//                String stringValue = null;
//                if(cell != null){
//                    stringValue = getCellValue(cell);
//                }
//                if (map[0].containsKey(rowNum + "," + colNum)) {
//                    String pointString = (String)map[0].get(rowNum + "," + colNum);
//                    int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
//                    int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
//                    int rowSpan = bottomeRow - rowNum + 1;
//                    int colSpan = bottomeCol - colNum + 1;
//                    if(map[2].containsKey(rowNum + "," + colNum)){
//                        rowSpan = rowSpan-(Integer)map[2].get(rowNum + "," + colNum);
////                        System.out.println(rowNum + "," + colNum+"减少rows:"+map[2].get(rowNum + "," + colNum)+"，剩余"+rowSpan+";当前值："+stringValue);
//                    }
//                    sb.append("<td rowspan= '" + rowSpan + "' colspan= '"+ colSpan + "' ");
//                    if(map.length>3 && map[3].containsKey(rowNum + "," + colNum)){
//                        //此类数据首行被隐藏，value为空，需使用其他方式获取值
//                        stringValue = getMergedRegionValue(sheet,rowNum,colNum);
//                    }
//                } else if (map[1].containsKey(rowNum + "," + colNum)) {
//                    map[1].remove(rowNum + "," + colNum);
//                    continue;
//                } else {
//                    sb.append("<td ");
//                }
//
//                //判断是否需要样式
//                if(isWithStyle){
//                    if(cell !=null ){
//                        dealExcelStyle(wb, sheet, cell, sb);//处理单元格样式
//                    }
//                    
//                }
//
//                sb.append(">");
//                if(sheetIndexPicMap!=null && sheetIndexPicMap.containsKey(imageRowNum)) sb.append(imageHtml);
//                if (stringValue == null || "".equals(stringValue.trim())) {
//                    sb.append("   ");
//                } else {
//                    // 将ascii码为160的空格转换为html下的空格（ ）
//                    sb.append(stringValue.replace(String.valueOf((char) 160)," "));
//                }
//                sb.append("</td>");
//            }
//            sb.append("</tr>");
//        }
//        sb.append("</table>");
//        return sb.toString();
//    }

            /**
      * 分析excel表格，记录合并单元格相关的参数，用于之后html页面元素的合并操作
      * @param sheet
      * @return
      */
            @SuppressWarnings({ "unchecked", "rawtypes" })
    private static Map<String, Object>[] getRowSpanColSpanMap(Sheet sheet) {
        Map<String, String> map0 = new HashMap<String, String>();//保存合并单元格的对应起始和截止单元格
        Map<String, String> map1 = new HashMap<String, String>();//保存被合并的那些单元格
        Map<String, Integer> map2 = new HashMap<String, Integer>();//记录被隐藏的单元格个数
        Map<String, String> map3 = new HashMap<String, String>();//记录合并了单元格，但是合并的首行被隐藏的情况
        int mergedNum = sheet.getNumMergedRegions();
        CellRangeAddress range = null;
        Row row = null;
        for (int i = 0; i < mergedNum; i++) {
            range = sheet.getMergedRegion(i);
            int topRow = range.getFirstRow();
            int topCol = range.getFirstColumn();
            int bottomRow = range.getLastRow();
            int bottomCol = range.getLastColumn();
            /**
                          * 此类数据为合并了单元格的数据
                          * 1.处理隐藏（只处理行隐藏，列隐藏poi已经处理）
                          */
            if(topRow!=bottomRow){
                int zeroRoleNum = 0;
                int tempRow = topRow;
                for(int j=topRow;j<=bottomRow;j++){
                    row = sheet.getRow(j);
                    if(row.getZeroHeight() || row.getHeight()==0){
                        if(j==tempRow){
                            //首行就进行隐藏，将rowTop向后移
                            tempRow++;
                            continue;//由于top下移，后面计算rowSpan时会扣除移走的列，所以不必增加zeroRoleNum;
                        }
                        zeroRoleNum++;
                    }
                 }
                 if(tempRow!=topRow){
                     map3.put(tempRow + "," + topCol,topRow + "," + topCol);
                     topRow = tempRow;
                 }
                if(zeroRoleNum!=0) map2.put(topRow + "," + topCol,zeroRoleNum);
            }
            map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
            int tempRow = topRow;
            while (tempRow <= bottomRow) {
                int tempCol = topCol;
                while (tempCol <= bottomCol) {
                    map1.put(tempRow + "," + tempCol, topRow + "," + topCol);
                    tempCol++;
                }
                tempRow++;
            }
            map1.remove(topRow + "," + topCol);
        }
        Map[] map = { map0, map1 ,map2,map3};
        System.err.println(map0);
        return map;
    }


    /**
      * 获取合并单元格的值
      * @param sheet
      * @param row
      * @param column
      * @return
      */
            public static String getMergedRegionValue(Sheet sheet, int row, int column){
        int sheetMergeCount = sheet.getNumMergedRegions();
        for(int i = 0 ; i < sheetMergeCount ; i++){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if(row >= firstRow && row <= lastRow){

                if(column >= firstColumn && column <= lastColumn){
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);

                    return getCellValue(fCell) ;
                }
            }
        }
        return null ;
    }
    /**
      * 获取表格单元格Cell内容
      * @param cell
      * @return
      */
            private static String getCellValue(Cell cell) {
        String result = new String();
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:// 数字类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                    SimpleDateFormat sdf = null;
                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                    }
                    Date date = cell.getDateCellValue();
                    result = sdf.format(date);
                } else if (cell.getCellStyle().getDataFormat() == 58) {
                    // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    double value = cell.getNumericCellValue();
                    Date date = org.apache.poi.ss.usermodel.DateUtil
                            .getJavaDate(value);
                    result = sdf.format(date);
                } else {
                    double value = cell.getNumericCellValue();
                    CellStyle style = cell.getCellStyle();
                    DecimalFormat format = new DecimalFormat();
                    String temp = style.getDataFormatString();
                    // 单元格设置成常规
                    if (temp.equals("General")) {
                        format.applyPattern("#");
                    }
                    result = format.format(value);
                }
                break;
            case Cell.CELL_TYPE_STRING:// String类型
                result = cell.getRichStringCellValue().toString();
                break;
            case Cell.CELL_TYPE_BLANK:
                result = "";
                break;
            default:
                result = "";
                break;
        }
        return result;
    }

    /**
      * 处理表格样式
      * @param wb
      * @param sheet
      * @param cell
      * @param sb
      */
            private static void dealExcelStyle(Workbook wb, Sheet sheet, Cell cell, StringBuffer sb){
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle != null) {
            short alignment = cellStyle.getAlignment();
            sb.append("align='" + convertAlignToHtml(alignment) + "' ");//单元格内容的水平对齐方式
            short verticalAlignment = cellStyle.getVerticalAlignment();
            sb.append("valign='"+ convertVerticalAlignToHtml(verticalAlignment)+ "' ");//单元格中内容的垂直排列方式

            if (wb instanceof XSSFWorkbook) {

                XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont();
                short boldWeight = xf.getBoldweight();
                sb.append("style='");
                sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
                sb.append("font-size: " + xf.getFontHeight() / 2 + "%;"); // 字体大小

                int topRow = cell.getRowIndex(),topColumn = cell.getColumnIndex();
                if(map[0].containsKey(topRow+","+topColumn)){//该单元格为合并单元格，宽度需要获取所有单元格宽度后合并
                    String value = (String)map[0].get(topRow+","+topColumn);
                    String[] ary = value.split(",");
                    int bottomColumn = Integer.parseInt(ary[1]);
                    if(topColumn!=bottomColumn){//合并列，需要计算相应宽度
                        int columnWidth = 0;
                        for(int i=topColumn;i<=bottomColumn;i++){
                            columnWidth += sheet.getColumnWidth(i);
                        }
                        sb.append("width:" + columnWidth/256*xf.getFontHeight()/20 + "pt;");
                    }else{
                        int columnWidth = sheet.getColumnWidth(cell.getColumnIndex()) ;
                        sb.append("width:" + columnWidth/256*xf.getFontHeight()/20 + "pt;");
                    }
                }else{
                    int columnWidth = sheet.getColumnWidth(cell.getColumnIndex()) ;
                    sb.append("width:" + columnWidth/256*xf.getFontHeight()/20 + "pt;");
                }

                XSSFColor xc = xf.getXSSFColor();
                if (xc != null && !"".equals(xc)) {
                    sb.append("color:#" + xc.getARGBHex().substring(2) + ";"); // 字体颜色
                }

                XSSFColor bgColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
                if (bgColor != null && !"".equals(bgColor)) {
                    sb.append("background-color:#" + bgColor.getARGBHex().substring(2) + ";"); // 背景颜色
                }
                sb.append("border:solid #000000 1px;");
//                sb.append(getBorderStyle(0,cellStyle.getBorderTop(), ((XSSFCellStyle) cellStyle).getTopBorderXSSFColor()));
//                sb.append(getBorderStyle(1,cellStyle.getBorderRight(), ((XSSFCellStyle) cellStyle).getRightBorderXSSFColor()));
//                sb.append(getBorderStyle(2,cellStyle.getBorderBottom(), ((XSSFCellStyle) cellStyle).getBottomBorderXSSFColor()));
//                sb.append(getBorderStyle(3,cellStyle.getBorderLeft(), ((XSSFCellStyle) cellStyle).getLeftBorderXSSFColor()));
            }else if(wb instanceof HSSFWorkbook){
                HSSFFont hf = ((HSSFCellStyle) cellStyle).getFont(wb);
                short boldWeight = hf.getBoldweight();
                short fontColor = hf.getColor();
                sb.append("style='");

                HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette(); // 类HSSFPalette用于求的颜色的国际标准形式
                HSSFColor hc = palette.getColor(fontColor);
                sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
                sb.append("font-size: " + hf.getFontHeight() / 2 + "%;"); // 字体大小
                String fontColorStr = convertToStardColor(hc);
                if (fontColorStr != null && !"".equals(fontColorStr.trim())) {
                    sb.append("color:" + fontColorStr + ";"); // 字体颜色
                }

                int topRow = cell.getRowIndex(),topColumn = cell.getColumnIndex();
                if(map[0].containsKey(topRow+","+topColumn)){//该单元格为合并单元格，宽度需要获取所有单元格宽度后合并
                    String value = (String)map[0].get(topRow+","+topColumn);
                    String[] ary = value.split(",");
                    int bottomColumn = Integer.parseInt(ary[1]);
                    if(topColumn!=bottomColumn){//合并列，需要计算相应宽度
                        int columnWidth = 0;
                        for(int i=topColumn;i<=bottomColumn;i++){
                            columnWidth += sheet.getColumnWidth(i);
                        }
                        sb.append("width:" + columnWidth/256*hf.getFontHeight()/20 + "pt;");
                    }else{
                        int columnWidth = sheet.getColumnWidth(cell.getColumnIndex()) ;
                        sb.append("width:" + columnWidth/256*hf.getFontHeight()/20 + "pt;");
                    }
                }else{
                    int columnWidth = sheet.getColumnWidth(cell.getColumnIndex()) ;
                    sb.append("width:" + columnWidth/256*hf.getFontHeight()/20 + "pt;");
                }

                short bgColor = cellStyle.getFillForegroundColor();
                hc = palette.getColor(bgColor);
                String bgColorStr = convertToStardColor(hc);
                if (bgColorStr != null && !"".equals(bgColorStr.trim())) {
                    sb.append("background-color:" + bgColorStr + ";"); // 背景颜色
                }
                sb.append("border:solid #000000 1px;");
//                sb.append( getBorderStyle(palette,0,cellStyle.getBorderTop(),cellStyle.getTopBorderColor()));
//                sb.append( getBorderStyle(palette,1,cellStyle.getBorderRight(),cellStyle.getRightBorderColor()));
//                sb.append( getBorderStyle(palette,3,cellStyle.getBorderLeft(),cellStyle.getLeftBorderColor()));
//                sb.append( getBorderStyle(palette,2,cellStyle.getBorderBottom(),cellStyle.getBottomBorderColor()));
            }
            sb.append("' ");
        }
    }

    /**
      * 单元格内容的水平对齐方式
      * @param alignment
      * @return
      */
            private static String convertAlignToHtml(short alignment) {
        String align = "left";
        switch (alignment) {
            case CellStyle.ALIGN_LEFT:
                align = "left";
                break;
            case CellStyle.ALIGN_CENTER:
                align = "center";
                break;
            case CellStyle.ALIGN_RIGHT:
                align = "right";
                break;
            default:
                break;
        }
        return align;
    }

    /**
      * 单元格中内容的垂直排列方式
      * @param verticalAlignment
      * @return
      */
            private static String convertVerticalAlignToHtml(short verticalAlignment) {
        String valign = "middle";
        switch (verticalAlignment) {
            case CellStyle.VERTICAL_BOTTOM:
                valign = "bottom";
                break;
            case CellStyle.VERTICAL_CENTER:
                valign = "center";
                break;
            case CellStyle.VERTICAL_TOP:
                valign = "top";
                break;
            default:
                break;
        }
        return valign;
    }

    private static String convertToStardColor(HSSFColor hc) {
        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
            }
        }
        return sb.toString();
    }

    private static String fillWithZero(String str) {
        if (str != null && str.length() < 2) {
            return "0" + str;
        }
        return str;
    }

    static String[] bordesr={"border-top:","border-right:","border-bottom:","border-left:"};
    static String[] borderStyles={"solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid","solid","solid","solid","solid"};

    private static  String getBorderStyle(HSSFPalette palette , int b, short s, short t){
        if(s==0)return  bordesr[b]+borderStyles[s]+"#d0d7e5 1px;";
        String borderColorStr = convertToStardColor( palette.getColor(t));
        borderColorStr=borderColorStr==null|| borderColorStr.length()<1?"#000000":borderColorStr;
        return bordesr[b]+borderStyles[s]+borderColorStr+" 1px;";
    }

    private static  String getBorderStyle(int b,short s, XSSFColor xc){
        if(s==0)return  bordesr[b]+borderStyles[s]+"#d0d7e5 1px;";;
        if (xc != null && !"".equals(xc)) {
            String borderColorStr = xc.getARGBHex();//t.getARGBHex();
            borderColorStr=borderColorStr==null|| borderColorStr.length()<1?"#000000":borderColorStr.substring(2);
            return bordesr[b]+borderStyles[s]+borderColorStr+" 1px;";
        }
        return "";
    }

    private static void writeFile(String content, String path) throws IOException {
//        OutputStream os = null;
//        BufferedWriter bw = null;
//        try {
//            File file = new File(path);
//            os = new FileOutputStream(file);
//            bw = new BufferedWriter(new OutputStreamWriter(os,"GBK"));
//            bw.write(content);
//        } catch (FileNotFoundException fnfe) {
//            fnfe.printStackTrace();
//        } catch (IOException ioe) {
//            ioe.printStackTrace();
//        } finally {
//            try {
//                if (bw != null)
//                    bw.close();
//                if (os != null)
//                    os.close();
//            } catch (IOException ie) {
//                ie.printStackTrace();
//            }
//        }
        
        File file2 = new File(path);
        StringBuilder sb = new StringBuilder();
        file2.createNewFile();// 创建文件

        sb.append("<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=GBK\"><title>Html Test</title></head><body>");
        sb.append("<div>");
        sb.append(content);
        sb.append("</div>");
        sb.append("</body></html>");

        PrintStream printStream = new PrintStream(new FileOutputStream(file2));
        printStream.println(sb.toString());// 将字符串写入文件
    }


    /**
      * 获取Excel图片公共方法
      * @param sheetNum 当前sheet编号
      * @param sheet 当前sheet对象
      * @param workbook 工作簿对象
      * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
      */
            public static Map<String, PictureData> getSheetPictrues(int sheetNum, Sheet sheet, Workbook workbook) {
        if(workbook instanceof  HSSFWorkbook){
            return getSheetPictrues03(sheetNum, (HSSFSheet) sheet, (HSSFWorkbook) workbook);
        }else if(workbook instanceof  XSSFWorkbook){
            return getSheetPictrues07(sheetNum, (XSSFSheet) sheet, (XSSFWorkbook) workbook);
        }else{
            return null;
        }
    }
    
        //    public static Map<String, PictureData> getSheetPictrues(int sheetNum, Sheet sheet, Workbook workbook) {
//        if(workbook instanceof  HSSFWorkbook){
//            return getSheetPictrues03(sheetNum, (HSSFSheet) sheet, (HSSFWorkbook) workbook);
//        }else if(workbook instanceof  XSSFWorkbook){
//            return getSheetPictrues07(sheetNum, (XSSFSheet) sheet, (XSSFWorkbook) workbook);
//        }else{
//            return null;
//        }
//    }

            /**
      * 获取Excel2003图片
      * @param sheetNum 当前sheet编号
      * @param sheet 当前sheet对象
      * @param workbook 工作簿对象
      * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
      * @throws IOException
      */
//    private static Map<String, PictureData> getSheetPictrues03(int sheetNum,
//                                                              HSSFSheet sheet, HSSFWorkbook workbook) {
//
////        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
//        Map<String,  PictureForExcelToHtml> sheetIndexPicMap = new HashMap<String,  PictureForExcelToHtml>();
//        List<HSSFPictureData> pictures = workbook.getAllPictures();
//        if (pictures.size() != 0) {
//            for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
//                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
//                shape.getLineWidth();
//                if (shape instanceof HSSFPicture) {
//                    HSSFPicture pic = (HSSFPicture) shape;
//                    int pictureIndex = pic.getPictureIndex() - 1;
//                    HSSFPictureData picData = pictures.get(pictureIndex);
//                    String picIndex = String.valueOf(sheetNum) + "_"
//                            + String.valueOf(anchor.getRow1()) + "_"
//                            + String.valueOf(anchor.getCol1());
//    sheetIndexPicMap.put(picIndex, picData);
//                     PictureForExcelToHtml picture = new  PictureForExcelToHtml(anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
//                    System.out.println("anchor.getCol1()="+anchor.getCol1()+" anchor.getCol2()="+anchor.getCol2());
//                    System.out.println("anchor.getRow1()="+anchor.getRow1()+" anchor.getRow2()="+anchor.getRow2());
//                }
//            }
//            return sheetIndexPicMap;
//        } else {
//            return null;
//        }
//    }
            
            private static Map<String, PictureData> getSheetPictrues03(int sheetNum,
            HSSFSheet sheet, HSSFWorkbook workbook) {

        // Map<String, PictureData> sheetIndexPicMap = new HashMap<String,
        // PictureData>();
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
        List<HSSFPictureData> pictures = workbook.getAllPictures();
        if (pictures.size() != 0) {
            for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                shape.getLineWidth();
                if (shape instanceof HSSFPicture) {
                    HSSFPicture pic = (HSSFPicture) shape;
                    int pictureIndex = pic.getPictureIndex() - 1;
                    HSSFPictureData picData = pictures.get(pictureIndex);
                    String picIndex = String.valueOf(sheetNum) + "_"
                            + String.valueOf(anchor.getRow1()) + "_"
                            + String.valueOf(anchor.getCol1());
                     PictureForExcelToHtml picture = new  PictureForExcelToHtml(anchor.getRow1(),
                                                        anchor.getCol1(), anchor.getRow2(),
                                                        anchor.getCol2());
                    sheetIndexPicMap.put(picIndex, picData);
                    wangPicMap.put(picIndex, picture);
                    System.out.println("anchor.getCol1()=" + anchor.getCol1()
                                                        + " anchor.getCol2()=" + anchor.getCol2());
                    System.out.println("anchor.getRow1()=" + anchor.getRow1()
                                                        + " anchor.getRow2()=" + anchor.getRow2());
                }
            }
            return sheetIndexPicMap;
        } else {
            return null;
        }
    }

    /**
      * 获取Excel2007图片
      * @param sheetNum 当前sheet编号
      * @param sheet 当前sheet对象
      * @param workbook 工作簿对象
      * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
      */
            private static Map<String, PictureData> getSheetPictrues07(int sheetNum,XSSFSheet sheet, XSSFWorkbook workbook) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();

        for (POIXMLDocumentPart dr : sheet.getRelations()) {
            if (dr instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) dr;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture pic = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = pic.getPreferredSize();
                    CTMarker ctMarker = anchor.getFrom();
                    String picIndex = String.valueOf(sheetNum) + "_"
                            + ctMarker.getRow() + "_" + ctMarker.getCol();
                     PictureForExcelToHtml picture = new  PictureForExcelToHtml(anchor.getRow1(),
                                                        anchor.getCol1(), anchor.getRow2(),
                                                        anchor.getCol2());
                    sheetIndexPicMap.put(picIndex,pic.getPictureData());
                    wangPicMap.put(picIndex, picture);
                    System.out.println("anchor.getCol1()="+anchor.getCol1()+" anchor.getCol2()="+anchor.getCol2());
                    System.out.println("anchor.getRow1()="+anchor.getRow1()+" anchor.getRow2()="+anchor.getRow2());
                    System.out.println("ctMarker.getRow()="+ctMarker.getRow()+" ctMarker.getCol()="+ctMarker.getCol());
                }
            }
        }
        return sheetIndexPicMap;
    }

    
        //    private static Map<String, PictureData> getSheetPictrues07(int sheetNum,
//            XSSFSheet sheet, XSSFWorkbook workbook) {
//        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
//
//        for (POIXMLDocumentPart dr : sheet.getRelations()) {
//            if (dr instanceof XSSFDrawing) {
//                XSSFDrawing drawing = (XSSFDrawing) dr;
//                List<XSSFShape> shapes = drawing.getShapes();
//                for (XSSFShape shape : shapes) {
//                    XSSFPicture pic = (XSSFPicture) shape;
//                    XSSFClientAnchor anchor = pic.getPreferredSize();
//                    CTMarker ctMarker = anchor.getFrom();
//                    String picIndex = String.valueOf(sheetNum) + "_"
//                            + ctMarker.getRow() + "_" + ctMarker.getCol();
//                    sheetIndexPicMap.put(picIndex, pic.getPictureData());
//                    System.out.println("anchor.getCol1()=" + anchor.getCol1()
//                            + " anchor.getCol2()=" + anchor.getCol2());
//                    System.out.println("anchor.getRow1()=" + anchor.getRow1()
//                            + " anchor.getRow2()=" + anchor.getRow2());
//                    System.out.println("ctMarker.getRow()=" + ctMarker.getRow()
//                            + " ctMarker.getCol()=" + ctMarker.getCol());
//                }
//            }
//        }
//
//        return sheetIndexPicMap;
//    }
            
            public static void printImg(List<Map<String, PictureData>> sheetList) throws IOException {
        for (Map<String, PictureData> map : sheetList) {
            printImg(map);
        }
    }

    public static void printImg(Map<String, PictureData> map) throws IOException {
        Object key[] = map.keySet().toArray();
        for (int i = 0; i < map.size(); i++) {
            // 获取图片流
            PictureData pic = map.get(key[i]);
            // 获取图片索引
            String picName = key[i].toString();
            // 获取图片格式
            String ext = pic.suggestFileExtension();

            byte[] data = pic.getData();

//            FileOutputStream out = new FileOutputStream("D:\\pic" + picName + "." + ext);
            FileOutputStream out = new FileOutputStream("C:\\D_files\\poi\\html\\pic" + picName + ".jpg");
            out.write(data);
            out.close();
        }
    }

    public static void main(String[] args) {
//        excelWriteToHtml("/home/dev/bb.xlsx", "/home/dev/pics/cbv.html");
        excelWriteToHtml("C:\\D_files\\make_files\\CNBankKeys.xlsx", "C:\\D_files\\poi\\html\\CNBankKeys.html");
        
    }
}
