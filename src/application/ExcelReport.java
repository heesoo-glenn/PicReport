package application;

import java.io.File;
import java.util.List;

public interface ExcelReport {
	
	public void execute(File pictureDir, File outputDir);
	public void readExcel(int positionColNumber,int pictureNoColNumber, int contentColNumber, File inExcel);
	public List<Object> getDmgStateAndPictures();
	public void setDmgStateAndPictures(List<Object> multsheet);
	
}
