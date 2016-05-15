package org.kotemaru.tool.mstotxt;

import java.io.File;

public interface Processor {
	public void start(String fileName, File file) throws Exception;
}
