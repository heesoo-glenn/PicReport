package application;

import java.io.File;
import java.util.List;

public interface ExcelReport {
	
	public void execute(File pictureDir, File outputDir,File  inputExcel,String pivot1Column, String pivot2Column, String selectedPrintType, ProgressEventHandler progressEventHandler);
	public void readExcel(int positionColNumber,int pictureNoColNumber, int contentColNumber, File inExcel);
	public List<Object> getDmgStateAndPictures();
	public void setDmgStateAndPictures(List<Object> multsheet);
	public void setInfoBeforeExecution(File inPictureDir, File outputDir, File inExcel, String pivot1Column_,
			String pivot2Column_, String selectedPrintType, ProgressEventHandler progressEventHandler);
	
}
