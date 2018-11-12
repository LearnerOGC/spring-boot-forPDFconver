import java.io.*;

public class FileUtils {
	private static final String ENCODING = "UTF-8";// UTF-8

	public static String GetFileExt(String path) {
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
	
	public static void createDir(String dirPath) {
		File dir = new File(dirPath);
		if (!dir.exists()) {
			dir.mkdirs();
		}

	}

	public static void writeFile(String content, String path) {
		createHtmlDir(path);
		OutputStream os = null;
		BufferedWriter bw = null;
		try {
			File file = new File(path);
			//如果pdf保存路径不存在，则创建路径
			if(!file.exists()){
				file.mkdirs();
			}
			os = new FileOutputStream(file);
			bw = new BufferedWriter(new OutputStreamWriter(os, ENCODING));
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
