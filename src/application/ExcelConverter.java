package application;
import java.io.*;
import java.util.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

public class ExcelConverter {

	private File path = null;
	private ArrayList<File> inputFiles = new ArrayList<File>();
	private HSSFWorkbook workbookOld = null;
	private XSSFWorkbook workbookNew = null;
	private int lastColumn = 0;
	private HashMap<Integer, XSSFCellStyle> styleMap = null;

	private void getInputFiles() {
		String call = "getInputFiles ";
		if (this.path.isFile()) {
			if (this.path.getAbsolutePath().endsWith(".xls")
					&& !new File(this.path.getAbsolutePath() + "x").exists())
				this.inputFiles.add(this.path);
			else {
				System.out.println("Datei endet nicht mit .xls oder XLSX-Datei existiert bereits");
			}
		} else
			for (File f : this.path.listFiles(new FilenameFilter() {
				// anonyme innere Klasse

				@Override
				public boolean accept(File dir, String name) {
					if (name.endsWith(".xls"))
						return true;
					return false;
				}

			})) {
				if (!new File(f.getAbsoluteFile() + "x").exists()) {
					this.inputFiles.add(f);
				}
			}
		System.out.println(call + "Dateien gefunden: " + this.inputFiles.size());
		System.out.println(call + "abgeschlossen");
	}

	private HSSFWorkbook getWorkBook(File f) throws FileNotFoundException,IOException {
		System.out.println("getWorkBook lese " + f.getAbsolutePath());
		POIFSFileSystem fs = new POIFSFileSystem(new BufferedInputStream(
				new FileInputStream(f)));
		HSSFWorkbook workbook = new HSSFWorkbook(fs);
		System.out.println("getWorkBook abgeschlossen");
		return workbook;
	}

	private void transform() {
		String call = "transform ";
		System.out.println(call + "Workbook");
		XSSFSheet sheetNew;
		HSSFSheet sheetOld;
		this.workbookNew.setForceFormulaRecalculation(this.workbookOld.getForceFormulaRecalculation());
		// workbookNew.setHidden(workbookOld.isHidden()); //ST@03.05.2012 -
		// von Apache noch nicht implementiert
		this.workbookNew.setMissingCellPolicy(this.workbookOld.getMissingCellPolicy());

		for (int i = 0; i < this.workbookOld.getNumberOfSheets(); i++) {
			sheetOld = this.workbookOld.getSheetAt(i);
			sheetNew = this.workbookNew.getSheet(sheetOld.getSheetName());
			System.out.println(call + "Sheet erstellt: "+ sheetOld.getSheetName());
			sheetNew = this.workbookNew.createSheet(sheetOld.getSheetName());
			this.transform(sheetOld, sheetNew);
		}
		System.out.println(call + "Anzahl verwendeter Styles: "	+ this.styleMap.size());
		System.out.println(call + "abgeschlossen");
	}

	private void transform(HSSFSheet sheetOld, XSSFSheet sheetNew) {
		System.out.println("transform Sheet");

		sheetNew.setDisplayFormulas(sheetOld.isDisplayFormulas());
		sheetNew.setDisplayGridlines(sheetOld.isDisplayGridlines());
		sheetNew.setDisplayGuts(sheetOld.getDisplayGuts());
		sheetNew.setDisplayRowColHeadings(sheetOld.isDisplayRowColHeadings());
		sheetNew.setDisplayZeros(sheetOld.isDisplayZeros());
		sheetNew.setFitToPage(sheetOld.getFitToPage());
		sheetNew.setForceFormulaRecalculation(sheetOld
				.getForceFormulaRecalculation());
		sheetNew.setHorizontallyCenter(sheetOld.getHorizontallyCenter());
		sheetNew.setMargin(Sheet.BottomMargin,
				sheetOld.getMargin(Sheet.BottomMargin));
		sheetNew.setMargin(Sheet.FooterMargin,
				sheetOld.getMargin(Sheet.FooterMargin));
		sheetNew.setMargin(Sheet.HeaderMargin,
				sheetOld.getMargin(Sheet.HeaderMargin));
		sheetNew.setMargin(Sheet.LeftMargin,
				sheetOld.getMargin(Sheet.LeftMargin));
		sheetNew.setMargin(Sheet.RightMargin,
				sheetOld.getMargin(Sheet.RightMargin));
		sheetNew.setMargin(Sheet.TopMargin, sheetOld.getMargin(Sheet.TopMargin));
		sheetNew.setPrintGridlines(sheetNew.isPrintGridlines());
		sheetNew.setRightToLeft(sheetNew.isRightToLeft());
		sheetNew.setRowSumsBelow(sheetNew.getRowSumsBelow());
		sheetNew.setRowSumsRight(sheetNew.getRowSumsRight());
		sheetNew.setVerticallyCenter(sheetOld.getVerticallyCenter());

		XSSFRow rowNew;
		for (Row row : sheetOld) {
			rowNew = sheetNew.createRow(row.getRowNum());
			if (rowNew != null)
				this.transform((HSSFRow) row, rowNew);
		}

		for (int i = 0; i < this.lastColumn; i++) {
			sheetNew.setColumnWidth(i, sheetOld.getColumnWidth(i));
			sheetNew.setColumnHidden(i, sheetOld.isColumnHidden(i));
		}

		for (int i = 0; i < sheetOld.getNumMergedRegions(); i++) {
			CellRangeAddress merged = sheetOld.getMergedRegion(i);
			sheetNew.addMergedRegion(merged);
		}
	}

	private void transform(HSSFRow rowOld, XSSFRow rowNew) {
		XSSFCell cellNew;
		rowNew.setHeight(rowOld.getHeight());
		if (rowOld.getRowStyle() != null) {
			Integer hash = rowOld.getRowStyle().hashCode();
			if (!this.styleMap.containsKey(hash))
				this.transform(hash, rowOld.getRowStyle(),
						this.workbookNew.createCellStyle());
			rowNew.setRowStyle(this.styleMap.get(hash));
		}
		for (Cell cell : rowOld) {
			cellNew = rowNew.createCell(cell.getColumnIndex(),cell.getCellType());
			if (cellNew != null)
				this.transform((HSSFCell) cell, cellNew);
		}
		this.lastColumn = Math.max(this.lastColumn, rowOld.getLastCellNum());
	}

	private void transform(HSSFCell cellOld, XSSFCell cellNew) {
		cellNew.setCellComment(cellOld.getCellComment());

		Integer hash = cellOld.getCellStyle().hashCode();
		if (!this.styleMap.containsKey(hash)) {
			this.transform(hash, cellOld.getCellStyle(),
					this.workbookNew.createCellStyle());
		}
		cellNew.setCellStyle(this.styleMap.get(hash));

		switch (cellOld.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			cellNew.setCellValue(cellOld.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_ERROR:
			cellNew.setCellValue(cellOld.getErrorCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA:
			cellNew.setCellValue(cellOld.getCellFormula());
			break;
		case Cell.CELL_TYPE_NUMERIC:
			cellNew.setCellValue(cellOld.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_STRING:
			cellNew.setCellValue(cellOld.getStringCellValue());
			break;
		default:
			System.out.println("transform: Unbekannter Zellentyp "+ cellOld.getCellType());
		}
	}

	private void transform(Integer hash, HSSFCellStyle styleOld,
			XSSFCellStyle styleNew) {
		styleNew.setAlignment(styleOld.getAlignment());
		styleNew.setBorderBottom(styleOld.getBorderBottom());
		styleNew.setBorderLeft(styleOld.getBorderLeft());
		styleNew.setBorderRight(styleOld.getBorderRight());
		styleNew.setBorderTop(styleOld.getBorderTop());
		styleNew.setDataFormat(this.transform(styleOld.getDataFormat()));
		styleNew.setFillBackgroundColor(styleOld.getFillBackgroundColor());
		styleNew.setFillForegroundColor(styleOld.getFillForegroundColor());
		styleNew.setFillPattern(styleOld.getFillPattern());
		styleNew.setFont(this.transform(styleOld.getFont(this.workbookOld)));
		styleNew.setHidden(styleOld.getHidden());
		styleNew.setIndention(styleOld.getIndention());
		styleNew.setLocked(styleOld.getLocked());
		styleNew.setVerticalAlignment(styleOld.getVerticalAlignment());
		styleNew.setWrapText(styleOld.getWrapText());
		this.styleMap.put(hash, styleNew);
	}

	private short transform(short index) {
		DataFormat formatOld = this.workbookOld.createDataFormat();
		DataFormat formatNew = this.workbookNew.createDataFormat();
		return formatNew.getFormat(formatOld.getFormat(index));
	}

	private XSSFFont transform(HSSFFont fontOld) {
		XSSFFont fontNew = this.workbookNew.createFont();
		fontNew.setBoldweight(fontOld.getBoldweight());
		fontNew.setCharSet(fontOld.getCharSet());
		fontNew.setColor(fontOld.getColor());
		fontNew.setFontName(fontOld.getFontName());
		fontNew.setFontHeight(fontOld.getFontHeight());
		fontNew.setItalic(fontOld.getItalic());
		fontNew.setStrikeout(fontOld.getStrikeout());
		fontNew.setTypeOffset(fontOld.getTypeOffset());
		fontNew.setUnderline(fontOld.getUnderline());
		return fontNew;
	}
}