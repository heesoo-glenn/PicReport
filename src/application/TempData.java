package application;

import javafx.beans.property.SimpleStringProperty;

public class TempData {
	private SimpleStringProperty content;
	private SimpleStringProperty position;
	private SimpleStringProperty pictureNo;
	private SimpleStringProperty pictureFile;
	
	public TempData(String rowStr, String fileName){
		String[] rowStrArr = rowStr.split(",");
		
		this.content = new SimpleStringProperty(rowStrArr[0]);
		this.pictureNo = new SimpleStringProperty(rowStrArr[1]);
		this.position = new SimpleStringProperty(rowStrArr[2]);
		this.pictureFile = new SimpleStringProperty(fileName);
		
		System.out.println(content);
		System.out.println(pictureNo);
		System.out.println(position);
		System.out.println(pictureFile);
		
	}

	public String getContent() {
		return content.get();
	}

	public void setContent(String content) {
		this.content.set(content);
	}

	public String getPosition() {
		return position.get();
	}

	public void setPosition(String position) {
		this.position.set(position);
	}

	public String getPictureNo() {
		return pictureNo.get();
	}

	public void setPictureNo(String pictureNo) {
		this.pictureNo.set(pictureNo);
	}

	public String getPictureFile() {
		return pictureFile.get();
	}

	public void setPictureFile(String pictureFile) {
		this.pictureFile.set(pictureFile);
	}

}
