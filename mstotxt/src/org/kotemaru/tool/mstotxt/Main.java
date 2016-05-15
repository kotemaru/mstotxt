package org.kotemaru.tool.mstotxt;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class Main {

	public static void main(String[] args) throws Exception {
		Main main = new Main();
		main.start(args[0]);
	}

	private void start(String fileName) throws Exception {
		Excel excel = new Excel();
		File file = new File(fileName);
		if (file.isFile()) {
			excel.start(fileName, file);
		} else {
			List<File> files = getExcelFiles(file, new ArrayList<File>());
			int len = file.getCanonicalPath().length();
			for (File aFile : files) {
				String fname = aFile.getCanonicalPath().substring(len+1);
				excel.start(fname, aFile);
			}
		}
	}


	private List<File> getExcelFiles(File dir, List<File> list) {
		File[] files = dir.listFiles();
		for (File file : files) {
			String name = file.getName();
			//System.out.println(">>>name="+name);
			if (file.isDirectory()) {
				getExcelFiles(file, list);
			} else if (name.endsWith(".xls") || name.endsWith(".xlsx")) {
				//System.out.println(">>>"+file);
				list.add(file);
			}
		}
		return list;
	}
}
