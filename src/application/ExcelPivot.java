package application;

import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STAxis;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STItemType;
public class ExcelPivot {  
	public void print_pivot(String pivot1,String pivot2, File file_root, File outputDir) throws Exception{
                /* Read the input file that contains the data to pivot */
                
				FileInputStream input_document;
				
				try {
				
					input_document = new FileInputStream(file_root);
				    
	                /* Create a POI XSSFWorkbook Object from the input file */
	                XSSFWorkbook my_xlsx_workbook = new XSSFWorkbook(input_document); 
	                /* Read Data to be Pivoted - we have only one worksheet */
	                XSSFSheet sheet = my_xlsx_workbook.getSheetAt(0); 
	                XSSFSheet pivot_sheet=my_xlsx_workbook.createSheet("피벗테이블");
	                
	                /* Get the reference for Pivot Data */
	                CellReference p1=new CellReference(pivot1);
	                CellReference p2=new CellReference(pivot2);
	                AreaReference a=new AreaReference(p1,p2);


	                /* Find out where the Pivot Table needs to be placed */
	                CellReference b=new CellReference("B2");
	                /* Create Pivot Table */
	                XSSFPivotTable pivotTable = pivot_sheet.createPivotTable(a,b,sheet);
	                /* Add filters */
	                //pivotTable.addReportFilter(0);
	                
	                pivotTable.addRowLabel(0);
	                
	                pivotTable.addRowLabel(2);
	                
	                pivotTable.addRowLabel(4);

	                pivotTable.addRowLabel(10);
	                
	                pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 9,"합계:물 량");
	                pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 8,"합계:개 소");


	                CTPivotFields pFields = pivotTable.getCTPivotTableDefinition().getPivotFields();
	                pFields.getPivotFieldArray(0).setOutline(false);
	                pFields.getPivotFieldArray(2).setOutline(false);
	                pFields.getPivotFieldArray(4).setOutline(false);
	                
	                /* Write Pivot Table to File */
	                File outputExcelFile = new File(outputDir.getPath() + "\\" + "피벗테이블.xlsx");
	    			FileOutputStream out = new FileOutputStream(outputExcelFile);
	                my_xlsx_workbook.write(out);
	                input_document.close();
                
				} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
        }
}
