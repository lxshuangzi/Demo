package stu.liuxiang.hm;

import java.io.File;
import java.io.IOException;

import javax.xml.crypto.dsig.spec.XSLTTransformParameterSpec;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class HMTestDemo {
	public static void main(String[] args){
		HMTestDemo demo = new HMTestDemo();
		demo.getAllPapers("F:\\百度云同步盘\\WOW\\试卷\\");
	}
	private File[] getAllPapers(String filePath){
		File fileDir = new File(filePath);
		File[] papers = fileDir.listFiles();
		return papers;
	}
	private void processPaper(File paper){
		try {
			Workbook workbook = WorkbookFactory.create(paper);
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
}
