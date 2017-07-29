package application;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;


import org.apache.poi.hssf.usermodel.HSSFHeader;
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
import org.apache.poi.ss.usermodel.RichTextString;
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
	private List<Object> multSheet;
	private List<String> sheetName;
	private Workbook wb;
	
	public void readExcel(int contentColNo, int pictureNoColNo,int positionColNo, File  inputExcel) {
		List<Object> multSheet = new ArrayList<Object>();
		List<String> sheetName = new ArrayList<>();
				
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
			// ù���� sheet ���� �б�
			for (int i = 0; i < sheet_number-1; i++) {
				
				List<DmgStateAndPicture> dmgStateAndPictures = new ArrayList<DmgStateAndPicture>();
				
				for (Row row : wb.getSheetAt(i)) {
					 //�� �б� 
					 String cellValue = readCellAsString(row.getCell(contentColNo));
					 String cellValue2 = readCellAsString(row.getCell(pictureNoColNo));
					 String cellValue3 = readCellAsString(row.getCell((positionColNo)-3));
					 String cellValue4 = readCellAsString(row.getCell((pictureNoColNo)-3));;
					 String cellValue5 = readCellAsString(row.getCell((pictureNoColNo)-2));;
					 String cellValue6 = readCellAsString(row.getCell((pictureNoColNo)-4));;
					 
					 String celldata1 = cellValue.replaceAll("\\p{Z}", "");
					 String celldata2 = cellValue2.replaceAll("\\p{Z}", "");
					 String celldata_tmp = cellValue3.replaceAll("\\p{Z}", "");
					 String celldata_sup = cellValue4.replaceAll("\\p{Z}", "");
					 String celldata_unit = cellValue5.replaceAll("\\p{Z}", "");
					 String celldata_ea = cellValue6.replaceAll("\\p{Z}", "");
					 
					 if(celldata_tmp.startsWith("����")){
						 cellValue_3_tmp = readCellAsString(row.getCell((positionColNo)-1));;
					 }				 
					 
					 if(!celldata1.equalsIgnoreCase("null") && !celldata1.equalsIgnoreCase("") && !celldata1.startsWith("����") &&
						!celldata2.equalsIgnoreCase("null") && !celldata2.equalsIgnoreCase("") && !celldata2.startsWith("����") &&
						!celldata_sup.equalsIgnoreCase("null") && !celldata_sup.equalsIgnoreCase("") && !celldata_sup.startsWith("����") &&
						!celldata_unit.equalsIgnoreCase("null") && !celldata_unit.equalsIgnoreCase("") && !celldata_unit.startsWith("����")&&
						!celldata_ea.equalsIgnoreCase("null") && !celldata_ea.equalsIgnoreCase("") && !celldata_ea.startsWith("����")
						){
						 dmgStateAndPictures.add(new DmgStateAndPicture(cellValue_3_tmp, cellValue, cellValue2, celldata_sup, celldata_unit,celldata_ea));	
						 //cellValue = ����, cellValue2 = picNO?, cellValue_3_tmp = ��ġ, celldata_sup =���� , celldata_unit = ���� ,celldata_ea = ����
						 //(String position, String content, String pictureFileNameInExcel)
					 }
					 
					 
					 
				}

				multSheet.add(dmgStateAndPictures);
				sheetName.add(wb.getSheetName(i));
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
        this.sheetName = sheetName;
        
		return;       
	}
	
	@Override
	public void execute(File pictureDir, File outputDir) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		List<Object> multSheet = this.multSheet;
		List<String> sheetName = this.sheetName;
				
		//��Ʈ��Ÿ��
		Font fontBody = workbook.createFont();
		fontBody.setColor(HSSFColor.BLACK.index);
		fontBody.setFontHeight((short)220);
		fontBody.setFontName("����ü");

		//����Ÿ��
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
		
		Row rowTemp;
		XSSFCell cellTemp;

		try {
			//�÷� ������ ����
			//sheet.setColumnBreak(2);
			
			for (int j = 0; j < multSheet.size(); j++) {
				
				Object sheets = multSheet.get(j);
				String sheet_name = sheetName.get(j);
				XSSFSheet sheet = workbook.createSheet(sheet_name+"_����"); //��������
				
				//��������
				sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
				
				Header pageHeader = sheet.getHeader();	//�Ӹ���
				pageHeader.setCenter(HSSFHeader.font("�޸տ�ü", "Normal") +HSSFHeader.fontSize((short) 26) + "�� �� �� ��");
							
	            //��� row ����
				rowTemp = sheet.createRow(0);		
				
				List<DmgStateAndPicture> dmgStateAndPictureSheet = (List<DmgStateAndPicture>) sheets;
				int rowcount = 9;
				int st_pic = 0;
				
				for (int i = 0; i < dmgStateAndPictureSheet.size(); i++) {				
					//�ο� ������ ����
					sheet.setRowBreak(rowcount);
					rowcount =  rowcount + 10;
					DmgStateAndPicture dmgStatePic = (DmgStateAndPicture)dmgStateAndPictureSheet.get(i);
					String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
					String position = dmgStatePic.getPosition();
					String content = dmgStatePic.getContent();
					String supply = dmgStatePic.getSupply();
					String unit = dmgStatePic.getUnit();
					String ea = dmgStatePic.getEa();
					
					String basePath = pictureDir.getPath();
					File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// ��������
					InputStream pictureFIS = new FileInputStream(pictureFile); 	// InputStream�� ���� set		*FileNotFoundException -> try catch �־�ߵ�
					byte[] bytes = IOUtils.toByteArray(pictureFIS);				// �̹��� binary�� byte �迭�� ����
	
					CreationHelper helper = workbook.getCreationHelper();
					XSSFDrawing drawing = sheet.createDrawingPatriarch();	//�׸� �����̳�, �׸��� ������ insert�ϴ� ��
					ClientAnchor anchor = helper.createClientAnchor();	// �׸��� ���� ��ǥ�� �����ϱ� ���� ��ü
	
					rowTemp = sheet.createRow(st_pic);
					rowTemp.setHeight((short)500);
					
					for (int k = 0;  k < 10; k++) {
						Cell cells1 = rowTemp.createCell(k);
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
					anchor.setRow1(st_pic);	// => A�� ����, 1���� �Ʒ��� �׸��� ������� = �������� �ȴ�.
					anchor.setCol2(9);					
					st_pic = st_pic+1;
					anchor.setRow2(st_pic);	// => D�� ����, 5���� �Ʒ��� �׸��� �����ϴ� = ������ �ȴ�
									
					int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook�� �߰��ϰ�, �߰��� �׸��� id ������ �޾ƿ�
					
					XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// ��ǥ(anchor)�� �׸�id�� �׸��� insert��	
					//( ���� ) picture.resize();	// ���� �̹��� ũ��� resize�ϴ°�
					
					rowTemp = sheet.createRow(st_pic);
					rowTemp.setHeight((short)500);
					for (int k = 0; k < 10; k++) {
						Cell cells1 = rowTemp.createCell(k);
						if(k==0){
							cells1.setCellStyle(picL_style);
						}else if(k==9){
							cells1.setCellStyle(picR_style);
						}
					}
					st_pic = st_pic +1;
					
					//���Cell ����
					rowTemp = sheet.createRow(st_pic);
					rowTemp.setHeight((short)500);
					for (int k = 0; k < 10; k++) {
						Cell cells1 = rowTemp.createCell(k);
						if(k == 0){
							cells1.setCellValue("��  ġ");
							cells1.setCellStyle(textheader_style);
						}else if(k == 2){
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
					
					for (int k = 0; k < 10; k++) {
						Cell cells1 = rowTemp.createCell(k);
						if(k == 0){
							cells1.setCellValue("��  ��");
							cells1.setCellStyle(textheader_style);
						}else if(k == 2){
							cells1.setCellValue(content);
							cells1.setCellStyle(text_style);
						}else if(k == 6){
							cells1.setCellValue(unit+" / "+supply+" / "+ea);
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
					        5  //last column  (0-based)
					));
					sheet.addMergedRegion(new CellRangeAddress(
							st_pic, //first row (0-based)
							st_pic, //last row  (0-based)
					        6, //first column (0-based)
					        9  //last column  (0-based)
					));
					
					st_pic = st_pic+1;				
	            }
			
			}
			//�������� END
            File outputExcelFile = new File(outputDir.getPath() + "\\" + "��������.xlsx");
			FileOutputStream out = new FileOutputStream(outputExcelFile);
			workbook.write(out);
            out.close();
            
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

	
	//�� ������ ������ Ȯ���ϰ� ������ string ������ ��ȯ��. 
	private String readCellAsString(Cell cell) {
		 String valueStr = "";
		 
		 if(cell != null){
			 switch(cell.getCellType()){
				case Cell.CELL_TYPE_STRING :
					valueStr = cell.getStringCellValue();
					break;
				case Cell.CELL_TYPE_NUMERIC : // ��¥ �����̵� ���� �����̵� �� CELL_TYPE_NUMERIC���� �ν���.
					if(DateUtil.isCellDateFormatted(cell)){ // ��¥ ������ �������� ���,
						SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd", Locale.KOREA);
						String formattedStr = dateFormat.format(cell.getDateCellValue());
						valueStr = formattedStr;
						break;
					}else{ // �����ϰ� ���� �������� ���,
						Double numericCellValue = cell.getNumericCellValue();
						if(Math.floor(numericCellValue) == numericCellValue){ // �Ҽ��� ���ϸ� ���� ���� ������ ���� ���ٸ�,,
							valueStr = numericCellValue.intValue() + ""; // int������ �Ҽ��� ���� ������ String���� ������ ��´�.
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

}
