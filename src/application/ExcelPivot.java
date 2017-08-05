package application;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotFields;

public class ExcelPivot {
	public void make_pivot(XSSFWorkbook workbook,String sheet_name, String pivot1Column_,String pivot2Column_,Integer num) {
		XSSFSheet sheet_original = workbook.getSheetAt(num); 				
		XSSFSheet pivot_sheet=workbook.createSheet(sheet_name+"_피벗테이블");
        
        CellReference p1=new CellReference(pivot1Column_);
        CellReference p2=new CellReference(pivot2Column_);
        AreaReference a=new AreaReference(p1,p2);

        CellReference b=new CellReference("B2");
        XSSFPivotTable pivotTable = pivot_sheet.createPivotTable(a,b,sheet_original);

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
	}
}
