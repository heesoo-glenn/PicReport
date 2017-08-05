package application;

import java.io.File;

public class DmgStateAndPicture {
	private String position;
	private String content;
	private String pictureFileNameInExcel;
	private File pictureFile;
	private String supply;
	private String unit;
	private String ea;
	private int sheetnum;//시트넘버 쓰려고 추가한거
	
	public DmgStateAndPicture(String position, String content, String pictureFileNameInExcel, String supply, String unit, String ea,int sheetnum){
		this.position = position;
		this.content = content;
		this.pictureFileNameInExcel = pictureFileNameInExcel;
		this.supply = supply;
		this.unit = unit;
		this.ea = ea;
		this.sheetnum = sheetnum;
	}	
	
	
	
	public int getSheetnum() {
		return sheetnum;
	}
	public void setSheetnum(int sheetnum) {
		this.sheetnum = sheetnum;
	}
	public String getEa() {
		return ea;
	}
	public void setEa(String ea) {
		this.ea = ea;
	}
	public String getSupply() {
		return supply;
	}
	public void setSupply(String supply) {
		this.supply = supply;
	}
	public String getUnit() {
		return unit;
	}
	public void setUnit(String unit) {
		this.unit = unit;
	}
	public String getPosition() {
		return position;
	}
	public void setPosition(String position) {
		this.position = position;
	}
	public String getContent() {
		return content;
	}
	public void setContent(String content) {
		this.content = content;
	}
	public String getPictureFileNameInExcel() {
		return pictureFileNameInExcel;
	}
	public void setPictureFileNameInExcel(String pictureFileNameInExcel) {
		this.pictureFileNameInExcel = pictureFileNameInExcel;
	}
	public File getPictureFile() {
		return pictureFile;
	}
	public void setPictureFile(File pictureFile) {
		this.pictureFile = pictureFile;
	}


}
