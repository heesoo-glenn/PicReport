package application;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReportXSSF implements ExcelReport{
	
	private File inputExcel;
	private List<DmgStateAndPicture> dmgStateAndPictures;
	
	public void readExcel(int contentColNo, int pictureNoColNo,int positionColNo, File  inputExcel) {
		List<DmgStateAndPicture> dmgStateAndPictures = new ArrayList<DmgStateAndPicture>();

        File file = inputExcel;
        Workbook wb;

        try {
			String fileName = inputExcel.getName();
			int lastDot = fileName.lastIndexOf('.');
			String extension = fileName.substring(lastDot+1);
			
			if(extension.equals("xls")){
				wb = new HSSFWorkbook(new FileInputStream(file));
			}else{
				wb = new XSSFWorkbook(new FileInputStream(file));
			}
			
			 Cell cell = null;
			 String cellValue3 = "";   
			 // 첫번재 sheet 내용 읽기
			 for (Row row : wb.getSheetAt(0)) {

				 //셀 읽기 
				 String cellValue = readCellAsString(row.getCell(contentColNo));
				 String cellValue2 = readCellAsString(row.getCell(pictureNoColNo));
				 String cellValue_tmp = readCellAsString(row.getCell((positionColNo)-3));
				 
				 String celldata1 = cellValue.replaceAll("\\p{Z}", "");
				 String celldata2 = cellValue2.replaceAll("\\p{Z}", "");
				 String celldata_tmp = cellValue_tmp.replaceAll("\\p{Z}", "");
				 if(celldata_tmp.startsWith("구간")){
					 cellValue3 = readCellAsString(row.getCell((positionColNo)-1));;
				 }
				 
				 if(!celldata1.equalsIgnoreCase("null") && !celldata1.equalsIgnoreCase("") && !celldata1.startsWith("결함") &&
					!celldata2.equalsIgnoreCase("null") && !celldata2.equalsIgnoreCase("") && !celldata2.startsWith("사진") ){
					 dmgStateAndPictures.add(new DmgStateAndPicture(cellValue3, cellValue, cellValue2));	
					 //cellValue = 내용, cellValue2 = picNO?, cellValue3 = 위치
					 //(String position, String content, String pictureFileNameInExcel)
				 }
			 }
		}catch (Exception e) {
			e.printStackTrace();
		}
        
        this.dmgStateAndPictures = dmgStateAndPictures;
		return;       
	}
	
	@Override
	public void execute(File pictureDir, File outputDir) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		List<DmgStateAndPicture> dmgStateAndPictures = this.dmgStateAndPictures;
		
		XSSFSheet sheet = workbook.createSheet("sheet1");
		sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
		
		Header pageHeader = sheet.getHeader();	//머릿말
		pageHeader.setCenter(HSSFHeader.font("휴먼옛체", "Normal") +
                HSSFHeader.fontSize((short) 26) + "사 진 대 지");
		
		
		Row rowTemp;
		XSSFCell cellTemp;
		
		int st_pic = 0;
		try {
            //출력 row 생성
			rowTemp = sheet.createRow(0);			
						
			//폰트스타일
			Font fontBody = workbook.createFont();
			fontBody.setColor(HSSFColor.BLACK.index);
			fontBody.setFontHeight((short)220);
			fontBody.setFontName("굴림체");

			//셀스타일
			CellStyle textheader_style = workbook.createCellStyle();
			textheader_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
			textheader_style.setAlignment(CellStyle.ALIGN_CENTER);
			textheader_style.setFont(fontBody);
			textheader_style.setBorderTop(CellStyle.BORDER_MEDIUM);                      
			textheader_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
			textheader_style.setBorderRight(CellStyle.BORDER_MEDIUM);
			textheader_style.setBorderBottom(CellStyle.BORDER_MEDIUM);
			
			CellStyle text_style = workbook.createCellStyle();
			text_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
			text_style.setFont(fontBody);
			text_style.setIndention((short)1);
			text_style.setBorderTop(CellStyle.BORDER_MEDIUM);                      
			text_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
			text_style.setBorderRight(CellStyle.BORDER_MEDIUM);
			text_style.setBorderBottom(CellStyle.BORDER_MEDIUM);

			CellStyle picTop_style = workbook.createCellStyle();
			picTop_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
			picTop_style.setBorderTop(CellStyle.BORDER_MEDIUM);  
			picTop_style.setBorderRight(CellStyle.BORDER_MEDIUM);
			
			CellStyle picL_style = workbook.createCellStyle();
			picL_style.setBorderLeft(CellStyle.BORDER_MEDIUM);

			CellStyle picR_style = workbook.createCellStyle();
			picR_style.setBorderRight(CellStyle.BORDER_MEDIUM);
/*
			//출력 cell 생성
			for (int i = 0; i < 10; i++) {
				Cell cells1 = row.createCell(i);
				cells1.setCellValue("사 진 대 지");
				cells1.setCellStyle(header);
				sheet.setColumnWidth((short)i, (short)1900);
			}
			
			//셀병합 - 3개의 셀을 합칠경우 3개의 셀을 생성해야함.
			sheet.addMergedRegion(new CellRangeAddress(
			        0, //first row (0-based)
			        0, //last row  (0-based)
			        0, //first column (0-based)
			        9  //last column  (0-based)
			));	
			
			row.setHeight((short)900);   */

			//컬럼 페이지 설정
			//sheet.setColumnBreak(2);
			int rowcount = 9;
            for (int i = 0; i < dmgStateAndPictures.size(); i++) {				
				//로우 페이지 설정
				sheet.setRowBreak(rowcount);
				rowcount =  rowcount + 10;
				DmgStateAndPicture dmgStatePic = (DmgStateAndPicture)dmgStateAndPictures.get(i);
				String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
				String position = dmgStatePic.getPosition();
				String content = dmgStatePic.getContent();
				
				String basePath = pictureDir.getPath();
				File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// 파일정보
				InputStream pictureFIS = new FileInputStream(pictureFile); 	// InputStream에 파일 set		*FileNotFoundException -> try catch 있어야됨
				byte[] bytes = IOUtils.toByteArray(pictureFIS);				// 이미지 binary를 byte 배열에 담음

				CreationHelper helper = workbook.getCreationHelper();
				XSSFDrawing drawing = sheet.createDrawingPatriarch();	//그림 컨테이너, 그림을 실제로 insert하는 놈
				ClientAnchor anchor = helper.createClientAnchor();	// 그림을 넣을 좌표를 지정하기 위한 객체

				rowTemp = sheet.createRow(st_pic);
				rowTemp.setHeight((short)500);
				
				for (int j = 0; j < 10; j++) {
					Cell cells1 = rowTemp.createCell(j);
					cells1.setCellStyle(picTop_style);
					cells1.setCellStyle(picTop_style);
				}
				sheet.addMergedRegion(new CellRangeAddress(
						st_pic, //first row (0-based)
						st_pic, //last row  (0-based)
				        0, //first column (0-based)
				        9  //last column  (0-based)
				));	
				st_pic = st_pic +1;
				
				rowTemp = sheet.createRow(st_pic);
				rowTemp.setHeight((short)4700);
				rowTemp.createCell(0).setCellStyle(picL_style);
				rowTemp.createCell(9).setCellStyle(picR_style);
				
				anchor.setCol1(1);	
				anchor.setRow1(st_pic);	// => A의 우측, 1행의 아래가 그림의 좌측상단 = 시작점이 된다.
				anchor.setCol2(9);					
				st_pic = st_pic+1;
				anchor.setRow2(st_pic);	// => D의 우측, 5행의 아래가 그림의 우측하단 = 끝점이 된다
								
				int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook에 추가하고, 추가한 그림의 id 같은걸 받아옴
				
				XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// 좌표(anchor)와 그림id로 그림을 insert함	
				//( 참고 ) picture.resize();	// 원래 이미지 크기로 resize하는것
				
				rowTemp = sheet.createRow(st_pic);
				rowTemp.setHeight((short)500);
				for (int j = 0; j < 10; j++) {
					Cell cells1 = rowTemp.createCell(j);
					if(j==0){
						cells1.setCellStyle(picL_style);
					}else if(j==9){
						cells1.setCellStyle(picR_style);
					}
				}
				st_pic = st_pic +1;
				
				//출력Cell 생성
				rowTemp = sheet.createRow(st_pic);
				rowTemp.setHeight((short)500);
				for (int j = 0; j < 10; j++) {
					Cell cells1 = rowTemp.createCell(j);
					if(j == 0){
						cells1.setCellValue("위  치");
						cells1.setCellStyle(textheader_style);
					}else if(j == 2){
						cells1.setCellValue(position);
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}
				}
				
				sheet.addMergedRegion(new CellRangeAddress(
						st_pic, //first row (0-based)
						st_pic, //last row  (0-based)
				        0, //first column (0-based)
				        1  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						st_pic, //first row (0-based)
						st_pic, //last row  (0-based)
				        2, //first column (0-based)
				        9  //last column  (0-based)
				));
				
				st_pic = st_pic+1;	
				rowTemp = sheet.createRow(st_pic);
				rowTemp.setHeight((short)500);
				
				for (int j = 0; j < 10; j++) {
					Cell cells1 = rowTemp.createCell(j);
					if(j == 0){
						cells1.setCellValue("내  용");
						cells1.setCellStyle(textheader_style);
					}else if(j == 2){
						cells1.setCellValue(content);
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}					
				}
				
				sheet.addMergedRegion(new CellRangeAddress(
						st_pic, //first row (0-based)
						st_pic, //last row  (0-based)
				        0, //first column (0-based)
				        1  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						st_pic, //first row (0-based)
						st_pic, //last row  (0-based)
				        2, //first column (0-based)
				        9  //last column  (0-based)
				));

				st_pic = st_pic+1;				
            }

            File outputExcelFile = new File(outputDir.getPath() + "\\" + "사진대지.xlsx");
			FileOutputStream out = new FileOutputStream(outputExcelFile);
			workbook.write(out);
            out.close();
            
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}

	@Override
	public List<DmgStateAndPicture> getDmgStateAndPictures() {
		return dmgStateAndPictures;
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
				default:
					break;
			}
		}
		return valueStr;		
	}

	@Override
	public void setDmgStateAndPictures(List<DmgStateAndPicture> dmgStateAndPictures) {
		this.dmgStateAndPictures = dmgStateAndPictures;
	}
	
	

}
