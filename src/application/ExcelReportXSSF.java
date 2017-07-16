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
			 // ù���� sheet ���� �б�
			 for (Row row : wb.getSheetAt(0)) {

				 //�� �б� 
				 String cellValue = readCellAsString(row.getCell(contentColNo));
				 String cellValue2 = readCellAsString(row.getCell(pictureNoColNo));
				 String cellValue_tmp = readCellAsString(row.getCell((positionColNo)-3));
				 
				 String celldata1 = cellValue.replaceAll("\\p{Z}", "");
				 String celldata2 = cellValue2.replaceAll("\\p{Z}", "");
				 String celldata_tmp = cellValue_tmp.replaceAll("\\p{Z}", "");
				 if(celldata_tmp.startsWith("����")){
					 cellValue3 = readCellAsString(row.getCell((positionColNo)-1));;
				 }
				 
				 if(!celldata1.equalsIgnoreCase("null") && !celldata1.equalsIgnoreCase("") && !celldata1.startsWith("����") &&
					!celldata2.equalsIgnoreCase("null") && !celldata2.equalsIgnoreCase("") && !celldata2.startsWith("����") ){
					 dmgStateAndPictures.add(new DmgStateAndPicture(cellValue3, cellValue, cellValue2));	
					 //cellValue = ����, cellValue2 = picNO?, cellValue3 = ��ġ
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
		
		Header pageHeader = sheet.getHeader();	//�Ӹ���
		pageHeader.setCenter(HSSFHeader.font("�޸տ�ü", "Normal") +
                HSSFHeader.fontSize((short) 26) + "�� �� �� ��");
		
		
		Row rowTemp;
		XSSFCell cellTemp;
		
		int st_pic = 0;
		try {
            //��� row ����
			rowTemp = sheet.createRow(0);			
						
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
/*
			//��� cell ����
			for (int i = 0; i < 10; i++) {
				Cell cells1 = row.createCell(i);
				cells1.setCellValue("�� �� �� ��");
				cells1.setCellStyle(header);
				sheet.setColumnWidth((short)i, (short)1900);
			}
			
			//������ - 3���� ���� ��ĥ��� 3���� ���� �����ؾ���.
			sheet.addMergedRegion(new CellRangeAddress(
			        0, //first row (0-based)
			        0, //last row  (0-based)
			        0, //first column (0-based)
			        9  //last column  (0-based)
			));	
			
			row.setHeight((short)900);   */

			//�÷� ������ ����
			//sheet.setColumnBreak(2);
			int rowcount = 9;
            for (int i = 0; i < dmgStateAndPictures.size(); i++) {				
				//�ο� ������ ����
				sheet.setRowBreak(rowcount);
				rowcount =  rowcount + 10;
				DmgStateAndPicture dmgStatePic = (DmgStateAndPicture)dmgStateAndPictures.get(i);
				String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
				String position = dmgStatePic.getPosition();
				String content = dmgStatePic.getContent();
				
				String basePath = pictureDir.getPath();
				File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// ��������
				InputStream pictureFIS = new FileInputStream(pictureFile); 	// InputStream�� ���� set		*FileNotFoundException -> try catch �־�ߵ�
				byte[] bytes = IOUtils.toByteArray(pictureFIS);				// �̹��� binary�� byte �迭�� ����

				CreationHelper helper = workbook.getCreationHelper();
				XSSFDrawing drawing = sheet.createDrawingPatriarch();	//�׸� �����̳�, �׸��� ������ insert�ϴ� ��
				ClientAnchor anchor = helper.createClientAnchor();	// �׸��� ���� ��ǥ�� �����ϱ� ���� ��ü

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
				anchor.setRow1(st_pic);	// => A�� ����, 1���� �Ʒ��� �׸��� ������� = �������� �ȴ�.
				anchor.setCol2(9);					
				st_pic = st_pic+1;
				anchor.setRow2(st_pic);	// => D�� ����, 5���� �Ʒ��� �׸��� �����ϴ� = ������ �ȴ�
								
				int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook�� �߰��ϰ�, �߰��� �׸��� id ������ �޾ƿ�
				
				XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// ��ǥ(anchor)�� �׸�id�� �׸��� insert��	
				//( ���� ) picture.resize();	// ���� �̹��� ũ��� resize�ϴ°�
				
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
				
				//���Cell ����
				rowTemp = sheet.createRow(st_pic);
				rowTemp.setHeight((short)500);
				for (int j = 0; j < 10; j++) {
					Cell cells1 = rowTemp.createCell(j);
					if(j == 0){
						cells1.setCellValue("��  ġ");
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
						cells1.setCellValue("��  ��");
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

            File outputExcelFile = new File(outputDir.getPath() + "\\" + "��������.xlsx");
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
