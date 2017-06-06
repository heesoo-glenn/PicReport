package application;

import java.io.File;

public class FormedPicture {
	private String position;
	private String content;
	private String pictureFileNameInExcel;
	private File pictureFile;
	
	
	public FormedPicture(String position, String content, String pictureFileNameInExcel){
		this.position = position;
		this.content = content;
		this.pictureFileNameInExcel = pictureFileNameInExcel;
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
