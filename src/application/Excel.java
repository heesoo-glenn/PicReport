package application;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	
	static HSSFRow row;
	static HSSFCell cell;
	
	private List read_data;
	private File inputExcel;
	
	//�������� ����
	public void execute(File pictureDir, File outputDir){
		
		
		List read_data = this.read_data;
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("sheet1");

		//�μ���� ����
		sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);

		Header pageHeader = sheet.getHeader();
		pageHeader.setCenter(HSSFHeader.font("�޸տ�ü", "Normal") +
                HSSFHeader.fontSize((short) 26) + "�� �� �� ��");
		
		int st_pic = 0;
		
		try {
            //��� row ����
			row = sheet.createRow(0);			
						
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
            for (int i = 0; i < read_data.size(); i++) {				
				//�ο� ������ ����
				sheet.setRowBreak(rowcount);
				rowcount =  rowcount + 10;
				String data = (String)read_data.get(i);
				String[] data_row = data.split(",");		
				
				String basePath = pictureDir.getPath();
				File test1Picture = new File(basePath+"\\"+data_row[1]+".jpg");			// ��������
				InputStream test1Stream = new FileInputStream(test1Picture); 	// InputStream�� ���� set		*FileNotFoundException -> try catch �־�ߵ�
				byte[] bytes = IOUtils.toByteArray(test1Stream);				// �̹��� binary�� byte �迭�� ����
				
	
				HSSFPatriarch drawing = sheet.createDrawingPatriarch();	//�׸� �����̳�, �׸��� ������ insert�ϴ� ��
				
				ClientAnchor anchor = new HSSFClientAnchor();	// �׸��� ���� ��ǥ�� �����ϱ� ���� ��ü
				
				/*
					col1, row1�� �׸��� �������, col2, row2�� �׸��� �����ϴ�
					������ ���ڸ� ������ܺ����� ��ǥ����̶� �����ϸ� ���
				*/
				row = sheet.createRow(st_pic);
				row.setHeight((short)500);
				
				for (int j = 0; j < 10; j++) {
					Cell cells1 = row.createCell(j);
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
				
				row = sheet.createRow(st_pic);
				row.setHeight((short)4700);
				row.createCell(0).setCellStyle(picL_style);
				row.createCell(9).setCellStyle(picR_style);
				
				anchor.setCol1(1);	
				anchor.setRow1(st_pic);	// => A�� ����, 1���� �Ʒ��� �׸��� ������� = �������� �ȴ�.
				anchor.setCol2(9);					
				st_pic = st_pic+1;
				anchor.setRow2(st_pic);	// => D�� ����, 5���� �Ʒ��� �׸��� �����ϴ� = ������ �ȴ�
								
				int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook�� �߰��ϰ�, �߰��� �׸��� id ������ �޾ƿ�
				
				HSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// ��ǥ(anchor)�� �׸�id�� �׸��� insert��	
				//( ���� ) picture.resize();	// ���� �̹��� ũ��� resize�ϴ°�
				
				row = sheet.createRow(st_pic);
				row.setHeight((short)500);
				for (int j = 0; j < 10; j++) {
					Cell cells1 = row.createCell(j);
					if(j==0){
						cells1.setCellStyle(picL_style);
					}else if(j==9){
						cells1.setCellStyle(picR_style);
					}
				}
				st_pic = st_pic +1;
				
				//���Cell ����
				row = sheet.createRow(st_pic);
				row.setHeight((short)500);
				for (int j = 0; j < 10; j++) {
					Cell cells1 = row.createCell(j);
					if(j == 0){
						cells1.setCellValue("��  ġ");
						cells1.setCellStyle(textheader_style);
					}else if(j == 2){
						cells1.setCellValue(data_row[2]);
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
				row = sheet.createRow(st_pic);
				row.setHeight((short)500);
				
				for (int j = 0; j < 10; j++) {
					Cell cells1 = row.createCell(j);
					if(j == 0){
						cells1.setCellValue("��  ��");
						cells1.setCellStyle(textheader_style);
					}else if(j == 2){
						cells1.setCellValue(data_row[0]);
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
			
            
            File outputExcelFile = new File(outputDir.getPath() + "\\" + "��������.xls");
			FileOutputStream out = new FileOutputStream(outputExcelFile);
			workbook.write(out);
            out.close();
            
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
		
	}

	//���� �б� ����ȣ 1,2,3 �Է��ʿ�
	@SuppressWarnings("deprecation")
	public void excelRead(int column_num1, int column_num2,int column_num3, File  inputExcel) {

		List<String> data_row = new ArrayList<String>();
		
		// ��������
        File file = inputExcel;

        // ���� ���� ����
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
				 String cellValue = excel_cellData(row.getCell(column_num1));
				 String cellValue2 = excel_cellData(row.getCell(column_num2));
				 
				 String cellValue_tmp = excel_cellData(row.getCell((column_num3)-3));
				 
				 String celldata1 = cellValue.replaceAll("\\p{Z}", "");
				 String celldata2 = cellValue2.replaceAll("\\p{Z}", "");
				 String celldata_tmp = cellValue_tmp.replaceAll("\\p{Z}", "");
				 if(celldata_tmp.startsWith("����")){
					 cellValue3 = excel_cellData(row.getCell((column_num3)-1));;
				 }
				 
				 if(!celldata1.equalsIgnoreCase("null") && !celldata1.equalsIgnoreCase("") && !celldata1.startsWith("����") &&
					!celldata2.equalsIgnoreCase("null") && !celldata2.equalsIgnoreCase("") && !celldata2.startsWith("����") ){
					 data_row.add(cellValue+","+cellValue2+","+ cellValue3);				 
				 }
			 }
		}catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
        this.read_data = data_row;
		return;       
	}
	
	//�� ������ ������ Ȯ���ϰ� ������ string ������ ��ȯ��. 
	public String excel_cellData(Cell cell) {
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

	public List getRead_data() {
		return read_data;
	}

	public void setRead_data(List read_data) {
		this.read_data = read_data;
	}

	public File getInputExcel() {
		return inputExcel;
	}

	public void setInputExcel(File inputExcel) {
		this.inputExcel = inputExcel;
	}
	
	//excelRead() ������ �����ؾ� �Ѵ�.
	public void createPivotTableOn(HSSFWorkbook workbook){
		HSSFSheet pivotSheet = workbook.createSheet("pivotSheet");
		//XSSF �ۿ� Pivot Table ������ �������� �ʴ´� ...
		//XSSFSheet sheet = my_xlsx_workbook.getSheetAt(0); 
        /* Get the reference for Pivot Data */
        //AreaReference a=new AreaReference("A1:C51");
        /* Find out where the Pivot Table needs to be placed */
        //CellReference b=new CellReference("I5");
        /* Create Pivot Table */
        //XSSFPivotTable pivotTable = sheet.createPivotTable(a,b);
        /* Add filters */
        //pivotTable.addReportFilter(0);
        //pivotTable.addRowLabel(1);
        //pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2); 
        /* Write Pivot Table to File */
		//List<String> data = this.read_data;
		
		
		
	}
	
}
