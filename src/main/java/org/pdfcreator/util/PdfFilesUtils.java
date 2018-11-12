package org.pdfcreator.util;

import org.apache.commons.lang3.StringUtils;
import org.pdfcreator.context.FileConstants;

import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PdfFilesUtils {
	public static String getFileNameWithOutSuffix(String path){
		if(!StringUtils.isBlank(path)){
			return path.substring(path.lastIndexOf("\\")+1,path.lastIndexOf("."));
		}else {
			return null;
		}

	}


	public static String getFileSuffix(String path) {
		String ext = null;
		int i = path.lastIndexOf('.');
		if (i > 0 && i < path.length() - 1) {
			ext = path.substring(i + 1).toLowerCase();
		}
		return ext;
	}

	public static void createHtmlDir(String path) {
		int i = path.lastIndexOf('/');
		String dirPath = "";
		if (i > 0 && i < path.length() - 1) {
			dirPath = path.substring(0, i).toLowerCase();
		}
		File dir = new File(dirPath);
		if (!dir.exists()) {
			dir.mkdirs();
		}

	}
	
	public static File createDir(String dirPath) {
		File dir = new File(dirPath);
		if (!dir.exists() || (dir.exists() && !dir.isDirectory())) {
			dir.mkdirs();
		}
		return dir;
	}

	public static File createParentDir(String dirPath) {
		File dir = new File(dirPath);
		File parent = dir.getParentFile();
		if (!parent.exists()) {
			parent.mkdirs();
		}
		return dir;
	}

	public static File createFile(String filePath) throws IOException{
		File file = new File(filePath);
		if(!file.exists()){
			file.createNewFile();
			return file;
		}
		return null;
	}

	public static void transformFontFamily(String sourcePath){
		File file = new File(sourcePath);
		StringBuilder sb = new StringBuilder();
		try {
			BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(file)));
			String line = null;
			while((line = reader.readLine()) != null){
				sb.append(line);
			}
			reader.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch (IOException ioe){
			ioe.printStackTrace();
		}
		String htmlFile = sb.toString();
		StringBuffer sbu = new StringBuffer();
		Pattern p = Pattern.compile("font-family:[^><;\"]*;");
		Matcher m = p.matcher(htmlFile);
		while(m.find()){
			String temp = m.group();
			String replaseStr = FileConstants.HTML_FONT_FAMILY;
			m.appendReplacement(sbu, replaseStr);
		}
		m.appendTail(sbu);
		writeFile(sbu.toString(), sourcePath);

	}


	public static void writeFile(String content, String path) {
		createHtmlDir(path);
		OutputStream os = null;
		BufferedWriter bw = null;
		try {
			File file = new File(path);
			File parent = file.getParentFile();
			//如果pdf保存路径不存在，则创建路径
			if(!parent.exists()){
				parent.mkdirs();
			}
			os = new FileOutputStream(file);
			bw = new BufferedWriter(new OutputStreamWriter(os, FileConstants.DEFAULT_ENCODING));
			bw.write(content);
		} catch (FileNotFoundException fnfe) {
			fnfe.printStackTrace();
		} catch (IOException ioe) {
			ioe.printStackTrace();
		} finally {
			try {
				if (bw != null)
					bw.close();
				if (os != null)
					os.close();
			} catch (IOException ie) {
				ie.printStackTrace();
			}
		}
	}

}
