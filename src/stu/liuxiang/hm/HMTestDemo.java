
package stu.liuxiang.hm;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.xml.crypto.dsig.spec.XSLTTransformParameterSpec;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.atp.AnalysisToolPak;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HMTestDemo {
	private String userName = "";
	public static void main(String[] args){
		HMTestDemo demo = new HMTestDemo();
		File[] papers = demo.getAllPapers("F:\\百度云同步盘\\WOW\\试卷\\");
		for(File paper : papers){
			demo.processPaper(paper);
		}
	}
	private File[] getAllPapers(String filePath){
		File fileDir = new File(filePath);
		File[] papers = fileDir.listFiles();
		return papers;
	}
	private void processPaper(File paper){
		List<String> answer = new ArrayList<String>();
		userName = paper.getName();
		System.out.println(paper.getName() + "跑分开始");
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(paper);
			XSSFSheet sheet = workbook.getSheetAt(0);
			int rowNum = 94;
			for(int i = 0;i < rowNum ; i++){
				String answerStr = "";
				if(sheet.getRow(i).getCell(6) != null){
					answerStr = sheet.getRow(i).getCell(6).getStringCellValue();
				}else{
//					System.out.println("row"+i + "==null");
				}
//				System.out.println("row" + i + answerStr);
				if(null != answerStr && !"".equals(answerStr)){
					answer.add(sheet.getRow(i).getCell(6).getStringCellValue());
				}
				
			}
			answer.remove(0);
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
		countScoreResult(answer);
		
	}
	private int countScoreResult(List<String> question){
		int totalCount = 0;
		List<String> countScore = new ArrayList<String>();
		List<String> wrongScore = new ArrayList<String>();
		getCountScore(countScore);
		System.out.println("question count:" + question.size());
		System.out.println("answer   count:" + countScore.size());
		if(question.size() != countScore.size()){
			System.out.println("questions count is not equal with answer count");
		}else{
			for(int i = 0;i<question.size();i++){
				if(question.get(i).equalsIgnoreCase(countScore.get(i)) || question.get(i).equalsIgnoreCase("F")){
					totalCount ++;
				}else{
					wrongScore.add(i +"=" + countScore.get(i));
				}
			}
		}
		String wrongStr = "";
		for(int i = 0 ; i<wrongScore.size();i++){
			wrongStr += wrongStr + ","+ wrongScore.get(i);
		}
		System.out.println("正确题目的数量为" + totalCount + "错误" + wrongStr);
//		System.out.println("正确题目的数量为" + totalCount);
		if(totalCount >=70){
			System.out.println(userName + "真是他吗个天才");
		}else if(totalCount<=60){
			System.out.println(userName + "妈的智障");
		}else{
			System.out.println(userName + "一般般啦");
		}
		return 0;
	}
	
	private void getCountScore(List<String> countScore){
		//1-5
		countScore.add("B");
		countScore.add("C");
		countScore.add("C");
		countScore.add("D");
		countScore.add("C");
		
		//6-10
		countScore.add("D");
		countScore.add("B");
		countScore.add("D");
		countScore.add("A");
		countScore.add("B");
		
		//11-15
		countScore.add("C");
		countScore.add("A");
		countScore.add("A");
		countScore.add("C");
		countScore.add("B");
	
		//16-20
		countScore.add("A");
		countScore.add("C");
		countScore.add("A");
		countScore.add("A");
		countScore.add("C");
		
		//21-25
		countScore.add("B");
		countScore.add("A");
		countScore.add("B");
		countScore.add("B");
		countScore.add("A");
		
		//26-30
		countScore.add("B");
		countScore.add("B");
		countScore.add("B");
		countScore.add("A");
		countScore.add("C");
		
		//31-35
		countScore.add("A");
		countScore.add("B");
		countScore.add("B");
		countScore.add("A");
		countScore.add("A");
		
		//36-40
		countScore.add("B");
		countScore.add("B");
		countScore.add("A");
		countScore.add("D");
		countScore.add("B");
		
		//41-45
		countScore.add("D");
		countScore.add("A");
		countScore.add("A");
		countScore.add("B");
		countScore.add("C");
		
		//46-50
		countScore.add("B");
		countScore.add("C");
		countScore.add("C");
		countScore.add("C");
		countScore.add("B");
		
		//51-55
		countScore.add("B");
		countScore.add("A");
		countScore.add("B");
		countScore.add("A");
		countScore.add("C");
		
		//56-60
		countScore.add("C");
		countScore.add("C");
		countScore.add("B");
		countScore.add("D");
		countScore.add("C");
		
		//61-65
		countScore.add("B");
		countScore.add("A");
		countScore.add("A");
		countScore.add("A");
		countScore.add("A");
		
		//66-70
		countScore.add("B");
		countScore.add("A");
		countScore.add("A");
		countScore.add("A");
		countScore.add("C");
		
		//71-74
		countScore.add("B");
		countScore.add("C");
		countScore.add("B");
		countScore.add("B");
		
	}
}
