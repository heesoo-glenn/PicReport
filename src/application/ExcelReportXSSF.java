package application;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javafx.collections.ObservableMap;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.stage.Stage;

public class ExcelReportXSSF implements ExcelReport, Runnable{
	
	private List<Object> multSheet;
	
	public void readExcel(int contentColNo, int pictureNoColNo,int positionColNo, File  inputExcel) {
		List<Object> multSheet = new ArrayList<Object>();
		Workbook wb;		
        File file = inputExcel;

        try {
			String fileName = inputExcel.getName();
			int lastDot = fileName.lastIndexOf('.');
			String extension = fileName.substring(lastDot+1);
			
			if(extension.equals("xls")){
				wb = new HSSFWorkbook(new FileInputStream(file));
			}else{
				wb = new XSSFWorkbook(new FileInputStream(file));
			}
			
			int sheet_number = wb.getNumberOfSheets();
		
			Cell cell = null;
			String cellValue_3_tmp = "";   
			// 첫번재 sheet 내용 읽기
			for (int i = 0; i < sheet_number-1; i++) {
				
				List<DmgStateAndPicture> dmgStateAndPictures = new ArrayList<DmgStateAndPicture>();
				
				for (Row row : wb.getSheetAt(i)) {
					 //셀 읽기 
					 String cellValue = readCellAsString(row.getCell(contentColNo)); //결함종류
					 String cellValue2 = readCellAsString(row.getCell(pictureNoColNo)); //사진번호
					 String cellValue3 = readCellAsString(row.getCell(positionColNo)); //경간
					 String cellValue3_1 = readCellAsString(row.getCell((positionColNo)+1)); //부재
					 
					 String cellValue4 = readCellAsString(row.getCell((pictureNoColNo)-3));; //물량
					 String cellValue5 = readCellAsString(row.getCell((pictureNoColNo)-2));; //단위
					 String cellValue6 = readCellAsString(row.getCell((pictureNoColNo)-4));; //개소
					 
					 String celldata1 = cellValue.replaceAll("\\p{Z}", "");
					 String celldata2 = cellValue2.replaceAll("\\p{Z}", "");
					 String celldata_tmp = cellValue3_1.replaceAll("\\p{Z}", "");
					 cellValue_3_tmp = celldata_tmp;
					 celldata_tmp = cellValue3.replaceAll("\\p{Z}", "");
					 cellValue_3_tmp = cellValue_3_tmp+ "("+celldata_tmp+")";
					 
					 String celldata_sup = cellValue4.replaceAll("\\p{Z}", "");
					 String celldata_unit = cellValue5.replaceAll("\\p{Z}", "");
					 String celldata_ea = cellValue6.replaceAll("\\p{Z}", "");
					 
					 
					 if(!celldata1.equalsIgnoreCase("null") && !celldata1.equalsIgnoreCase("") && !celldata1.startsWith("결함") &&
						!celldata2.equalsIgnoreCase("null") && !celldata2.equalsIgnoreCase("") && !celldata2.startsWith("사진") &&
						!celldata_sup.equalsIgnoreCase("null") && !celldata_sup.equalsIgnoreCase("") && !celldata_sup.startsWith("물량") &&
						!cellValue_3_tmp.equalsIgnoreCase("null") && !cellValue_3_tmp.equalsIgnoreCase("") && !cellValue_3_tmp.startsWith("부재") &&
						!celldata_unit.equalsIgnoreCase("null") && !celldata_unit.equalsIgnoreCase("") && !celldata_unit.startsWith("단위")&&
						!celldata_ea.equalsIgnoreCase("null") && !celldata_ea.equalsIgnoreCase("") && !celldata_ea.startsWith("개소")
						){
						 dmgStateAndPictures.add(new DmgStateAndPicture(cellValue_3_tmp, cellValue, cellValue2, celldata_sup, celldata_unit,celldata_ea,i+1));	
						 //cellValue = 내용, cellValue2 = picNO?, cellValue_3_tmp = 위치, celldata_sup =물량 , celldata_unit = 단위 ,celldata_ea = 개소
						 //(String position, String content, String pictureFileNameInExcel)
					 }
					 
					 
					 
				}

				multSheet.add(dmgStateAndPictures);
			}
		}catch (Exception e) {
			e.printStackTrace();
			StringWriter sw = new StringWriter();
			e.printStackTrace(new PrintWriter(sw));
			String exceptionAsString = sw.toString();

			ExceptionCheck exx = new ExceptionCheck();
			try {
				exx.ExceptionCall(exceptionAsString);
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
		}
        
        this.multSheet = multSheet;        
		return;       
	}
	
	@Override
	public void execute(File pictureDir, File outputDir,File  inputExcel, String pivot1Column_, String pivot2Column_,int pictureNoColumn_,
			String selectedPrintType, ProgressEventHandler progressEventHandler) {
		//XSSFWorkbook workbook = new XSSFWorkbook();
		List<Object> multSheet = this.multSheet;

		FileInputStream input_document;

		ExcelImage image_make = new ExcelImage();

		ExcelPivot pivots = new ExcelPivot();

		try {
			progressEventHandler.gettingStart();
			
			input_document = new FileInputStream(inputExcel);
			
			progressEventHandler.readInputExcel();
			XSSFWorkbook workbook = new XSSFWorkbook(input_document); 	
			Thread.sleep(2000);
			progressEventHandler.makeOutputExcelData();
			Thread.sleep(2000);
			for (int j = 0; j < multSheet.size(); j++) {
				
				Object sheets = multSheet.get(j);
				String sheet_name = workbook.getSheetName(j);
				XSSFSheet sheet = workbook.createSheet(sheet_name+"_사진"); //사진대지
				
				//사진대지
				Header pageHeader = sheet.getHeader();	//머릿말
				pageHeader.setCenter(HSSFHeader.font("휴먼옛체", "Normal") +HSSFHeader.fontSize((short) 26) + "사 진 대 지");
				
				switch (selectedPrintType) {//출력시 사진대지부분을 몃열로 출력할지 정하는부분
				case "1": //1열
					sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
					image_make.make_1(pictureDir, workbook, sheet, sheets, pictureNoColumn_);
					
					int data_st_pic1 = sheet.getLastRowNum();
					
					int dats1 = workbook.getSheetIndex(sheet_name+"_사진");
					
					workbook.setPrintArea(
							dats1, //sheet index
							0, //start column
							9, //end column
							0, //start row
							data_st_pic1 //end row
					);
					sheet.setDisplayGridlines(true);
				    sheet.setPrintGridlines(true);
					break;
				case "2": //2열
					sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
					image_make.make_2(pictureDir, workbook, sheet, sheets, pictureNoColumn_);
					
					int data_st_pic2 = sheet.getLastRowNum();	
					int dats2 = workbook.getSheetIndex(sheet_name+"_사진");
					
					workbook.setPrintArea(
							dats2, //sheet index
							0, //start column
							19, //end column
							0, //start row
							data_st_pic2 //end row
					);			
					sheet.setDisplayGridlines(true);
				    sheet.setPrintGridlines(true);
					
					break;
				case "3": //3열
					sheet.getPrintSetup().setPaperSize(PrintSetup.A3_PAPERSIZE);
					image_make.make_3(pictureDir, workbook, sheet, sheets, pictureNoColumn_);
					
					int data_st_pic3 = sheet.getLastRowNum();
					int dats3 = workbook.getSheetIndex(sheet_name+"_사진");
					
					workbook.setPrintArea(
							dats3, //sheet index
							0, //start column
							29, //end column
							0, //start row
							data_st_pic3 //end row
					);	
					sheet.setDisplayGridlines(true);
				    sheet.setPrintGridlines(true);
					
					break;
				case "4": //4열
					sheet.getPrintSetup().setPaperSize(PrintSetup.A3_PAPERSIZE);
					image_make.make_4(pictureDir, workbook, sheet, sheets, pictureNoColumn_);	
					
					int data_st_pic4 = sheet.getLastRowNum();
					int dats4 = workbook.getSheetIndex(sheet_name+"_사진");
					
					workbook.setPrintArea(
							dats4, //sheet index
							0, //start column
							39, //end column
							0, //start row
							data_st_pic4 //end row
					);	
					sheet.setDisplayGridlines(true);
				    sheet.setPrintGridlines(true);
					
					break;
				default:
					break;
				}
				//사진대지 END
				
				//피벗
				pivots.make_pivot(workbook, sheet_name, pivot1Column_, pivot2Column_, j);
				//피벗 END
			}
            
			File outputExcelFile = new File(outputDir.getPath() + "\\" + "사진대지"+selectedPrintType+"열.xlsx");
			FileOutputStream out = new FileOutputStream(outputExcelFile);
			workbook.write(out);
            out.close();
            progressEventHandler.endProgress();

		} catch (Exception e) {
			e.printStackTrace();
			StringWriter sw = new StringWriter();
			e.printStackTrace(new PrintWriter(sw));
			String exceptionAsString = sw.toString();

			ExceptionCheck exx = new ExceptionCheck();
			try {
				exx.ExceptionCall(exceptionAsString);
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
		}		
	}

	
	//셀 데이터 형식을 확인하고 내용을 string 형으로 변환함. 
	private String readCellAsString(Cell cell) {
		 String valueStr = "";
		 
		 if(cell != null){
			 switch(cell.getCellType()){
				case Cell.CELL_TYPE_STRING :
					valueStr = cell.getStringCellValue();
					break;
				case Cell.CELL_TYPE_NUMERIC : // 날짜 형식이든 숫자 형식이든 다 CELL_TYPE_NUMERIC으로 인식함.
					if(DateUtil.isCellDateFormatted(cell)){ // 날짜 유형의 데이터일 경우,
						SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd", Locale.KOREA);
						String formattedStr = dateFormat.format(cell.getDateCellValue());
						valueStr = formattedStr;
						break;
					}else{ // 순수하게 숫자 데이터일 경우,
						Double numericCellValue = cell.getNumericCellValue();
						if(Math.floor(numericCellValue) == numericCellValue){ // 소수점 이하를 버린 값이 원래의 값과 같다면,,
							valueStr = numericCellValue.intValue() + ""; // int형으로 소수점 이하 버리고 String으로 데이터 담는다.
						}else{
							valueStr = numericCellValue + "";
						}
						break;
					}
				case Cell.CELL_TYPE_BOOLEAN :
					valueStr = cell.getBooleanCellValue() + "";
					break;
				case Cell.CELL_TYPE_ERROR :
					valueStr = cell.getBooleanCellValue() + "";
					break;
				case Cell.CELL_TYPE_FORMULA :
					switch(cell.getCachedFormulaResultType()) {
		            case Cell.CELL_TYPE_NUMERIC:
		            	valueStr = String.format("%.2f",cell.getNumericCellValue()); 
		                break;
		            case Cell.CELL_TYPE_STRING:
		            	RichTextString data = cell.getRichStringCellValue();
		            	valueStr = data.toString();
		                break;
					}
					break;
				default:
					break;
			}
		}
		return valueStr;		
	}

	@Override
	public void setDmgStateAndPictures(List<Object> multsheets) {
		this.multSheet = multsheets;
	}
	
	@Override
	public List<Object> getDmgStateAndPictures() {
		return multSheet;
	}

	@Override
	public void run() {
		runnableExecute();
	}
	
	private void runnableExecute(){
		execute(this.pictureDir, this.outputDir,this.inputExcel, this.pivot1Column_, this.pivot2Column_, this.pictureNoColumn_,this.selectedPrintType, this.progressEventHandler);
	}

	File pictureDir;
	File outputDir;
	File  inputExcel;
	String pivot1Column_;
	String pivot2Column_;
	int pictureNoColumn_;	
	String selectedPrintType;
	ProgressEventHandler progressEventHandler;
	
	@Override
	public void setInfoBeforeExecution(File pictureDir, File outputDir,File  inputExcel, String pivot1Column_, String pivot2Column_,int pictureNoColumn_,
			String selectedPrintType, ProgressEventHandler progressEventHandler) {

	this.pictureDir = pictureDir;
	this.outputDir = outputDir;
	this.inputExcel = inputExcel;
	this.pivot1Column_ = pivot1Column_;
	this.pivot2Column_ = pivot2Column_;
	this.pictureNoColumn_ = pictureNoColumn_;
	this.selectedPrintType = selectedPrintType;
	this.progressEventHandler = progressEventHandler;
	}
	


}
