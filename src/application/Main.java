package application;
 
import java.io.File;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
/* //xlsx 파일 출력시 선언
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
*/

import javafx.application.Application;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.ObservableMap;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.RadioButton;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.ToggleGroup;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.concurrent.Task;; 

public class Main extends Application {
	
	static File inExcel;	// 입력 엑셀
	static File inPictureDir;	// 입력 그림 폴더

	@Override
	public void start(Stage primaryStage) {
		try {
			FXMLLoader loader = new FXMLLoader();
			loader.setLocation(getClass().getResource("/resources/Main.fxml"));
			ObservableMap<String, Object> mainFXMLNamespace =  loader.getNamespace();
			Scene scene = new Scene(loader.load());
			
			ExcelReport excel = new ExcelReportXSSF();
			Util util = new Util();
			
			//엑셀선택 버튼 START
			Button setExcelButton = (Button) mainFXMLNamespace.get("SetInputExcelButton");
			setExcelButton.setOnMouseClicked(e -> {
				setExcelButton.setStyle("-fx-background-color:#e6ccff; -fx-border-color:#52527a;");
				FileChooser fileChooser = new FileChooser();
				inExcel = (fileChooser.showOpenDialog(primaryStage));
				Label excelPathLabel = (Label)mainFXMLNamespace.get("ExcelPathLabel");
				excelPathLabel.setText(inExcel.getAbsolutePath());
				setExcelButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
				return;				
			});
			setExcelButton.setOnMouseEntered(e->{
				setExcelButton.setStyle("-fx-background-color:#e6ccff; -fx-border-color:#52527a;");
			});
			setExcelButton.setOnMouseExited(e ->{
				setExcelButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
			});
			//엑셀선택버튼 END
			
			
			//그림폴더 선택 버튼 START
			Button setPicDirButton = (Button) mainFXMLNamespace.get("SetPicDirButton");
			setPicDirButton.setOnMouseClicked(e->{
				setPicDirButton.setStyle("-fx-background-color:#e6ccff; -fx-border-color:#52527a;");
				DirectoryChooser dirChooser = new DirectoryChooser();
				inPictureDir = (dirChooser.showDialog(primaryStage));
				Label picDirPathLabel = (Label)mainFXMLNamespace.get("PicDirPathLabel");
				picDirPathLabel.setText(inPictureDir.getAbsolutePath());
				setPicDirButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
			});
			setPicDirButton.setOnMouseEntered(e->{
				setPicDirButton.setStyle("-fx-background-color:#e6ccff; -fx-border-color:#52527a;");
			});
			setPicDirButton.setOnMouseExited(e ->{
				setPicDirButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
			});
			//그림폴더 선택 버튼 END
			
			
			//미리보기 버튼 START
			Button previewButton = (Button) mainFXMLNamespace.get("PreviewButton");
			previewButton.setOnMouseClicked(e ->{
				previewButton.setStyle("-fx-background-color:#e6ccff; -fx-border-color:#52527a;");

				String positionColumn_ =  ( (TextField) mainFXMLNamespace.get("PositionColumnTextField") ).getText();
				String contentColumn_ =  ( (TextField) mainFXMLNamespace.get("ContentColumnTextField") ).getText();
				String pictureNoColumn_ =  ( (TextField) mainFXMLNamespace.get("PictureNoColumnTextField") ).getText();
				
				int positionColNo = util.decodeToDecimal(positionColumn_);
				int contentColNo = util.decodeToDecimal(contentColumn_);
				int pictureNoColNo = util.decodeToDecimal(pictureNoColumn_);

				excel.readExcel(contentColNo, pictureNoColNo, positionColNo, inExcel);
				
				List<Object> multSheets =  excel.getDmgStateAndPictures();
				ObservableList<DmgStateAndPicture> dataList = FXCollections.observableArrayList();

				TableView tv = (TableView) mainFXMLNamespace.get("PreviewTableView");
				
				ObservableList<TableColumn> colLi = tv.getColumns();
				
				TableColumn sheetCol = colLi.get(0);
				TableColumn positionCol = colLi.get(1);	// 위치 : 0
				TableColumn contentCol = colLi.get(2);	//사진번호 : 1
				TableColumn pictureNoCol = colLi.get(3);
				TableColumn pictureFile = colLi.get(4);

			
				for (int i = 0; i < multSheets.size(); i++) {
					
					Object sheets = multSheets.get(i);
					List<DmgStateAndPicture> dmgStateAndPictureSheet = (List<DmgStateAndPicture>) sheets;

					checkPictureFileIsExists(dmgStateAndPictureSheet); // 실제로 그림파일 폴더에 해당하는 파일명의 그림파일이 있는지 확인한다. 해당 파일의 fullname을 갖고온다.
					HashMap<String, List<DmgStateAndPicture>> dupObjs = getDSPsDuplicatedOnPictureNumber(dmgStateAndPictureSheet,Integer.toString(i+1)); 
					
					
					sheetCol.setCellValueFactory(new PropertyValueFactory<DmgStateAndPicture,String>("sheetnum"));
					positionCol.setCellValueFactory(new PropertyValueFactory<DmgStateAndPicture,String>("position"));
					pictureNoCol.setCellValueFactory(new PropertyValueFactory<DmgStateAndPicture,String>("pictureFileNameInExcel"));
					contentCol.setCellValueFactory(new PropertyValueFactory<DmgStateAndPicture,String>("content"));
					pictureFile.setCellValueFactory(new PropertyValueFactory<DmgStateAndPicture,String>("pictureFile"));
					for(DmgStateAndPicture dmgStatPic  : dmgStateAndPictureSheet){
						dataList.add(dmgStatPic);
					}
					
				}				
				tv.setItems(dataList);
				
				previewButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
			});
			previewButton.setOnMouseEntered(e->{
				previewButton.setStyle("-fx-background-color:#e6ccff; -fx-border-color:#52527a;");
			});
			previewButton.setOnMouseExited(e ->{
				previewButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
			});
			//미리보기 버튼 END
			
			//생성버튼 START
			Button executeButton = (Button) mainFXMLNamespace.get("ExecuteButton");
			executeButton.setOnMouseClicked(e ->{
				executeButton.setStyle("-fx-background-color:#e6ccff; -fx-border-color:#52527a;");
				
				Alert alert = new Alert(AlertType.INFORMATION);
				alert.setTitle("진행");
				alert.setHeaderText(null);
				alert.setContentText("출력 엑셀을 저장할 폴더를 선택해 주세요.");
				alert.showAndWait();

				DirectoryChooser dirChooser = new DirectoryChooser();
				File outputDir = dirChooser.showDialog(primaryStage);
				
				if(outputDir == null ) {return;}
				
				ToggleGroup outputTypeToggleGroup = (ToggleGroup)mainFXMLNamespace.get("OutputTypeToggleGroup");
				RadioButton selectedRB = (RadioButton) outputTypeToggleGroup.getSelectedToggle();
				String selectedOutputType = selectedRB.getUserData().toString(); //xls  또는 xlsx
				
				ExcelReport outExcel = null;
				if(selectedOutputType.equals("xls")){
					outExcel = excel;
				}else{
					outExcel = new ExcelReportXSSF();
					outExcel.setDmgStateAndPictures(excel.getDmgStateAndPictures());
				}

				if(outExcel.getDmgStateAndPictures() == null ) {
					//엑셀 컬럼알파벳을 번호로 변환
					String positionColumn_ =  ( (TextField) mainFXMLNamespace.get("PositionColumnTextField") ).getText();
					String contentColumn_ =  ( (TextField) mainFXMLNamespace.get("ContentColumnTextField") ).getText();
					String pictureNoColumn_ =  ( (TextField) mainFXMLNamespace.get("PictureNoColumnTextField") ).getText();
					int positionColNo = util.decodeToDecimal(positionColumn_);
					int pictureNoColNo = util.decodeToDecimal(pictureNoColumn_);
					int contentColNo = util.decodeToDecimal(contentColumn_);

					outExcel.readExcel(contentColNo, pictureNoColNo, positionColNo, inExcel);
					
				}
				outExcel.execute(inPictureDir, outputDir);

				executeButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
			});
			executeButton.setOnMouseEntered(e->{
				executeButton.setStyle("-fx-background-color:#e6ccff; -fx-border-color:#52527a;");
			});
			executeButton.setOnMouseExited(e ->{
				executeButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
			});
			//생성버튼 END
			
			Button pivotButton = (Button) mainFXMLNamespace.get("PivotButton");
			pivotButton.setOnMouseClicked(e ->{
				ExcelPivot pivot = new ExcelPivot();
				try {
					Alert alert = new Alert(AlertType.INFORMATION);
					alert.setTitle("진행");
					alert.setHeaderText(null);
					alert.setContentText("출력 엑셀을 저장할 폴더를 선택해 주세요.");
					alert.showAndWait();

					DirectoryChooser dirChooser = new DirectoryChooser();
					File outputDir = dirChooser.showDialog(primaryStage);

					if(outputDir == null ) {return;}
					
					String pivot1Column_ =  ( (TextField) mainFXMLNamespace.get("Pivot1NoColumnTextField") ).getText();
					String pivot2Column_ =  ( (TextField) mainFXMLNamespace.get("Pivot2NoColumnTextField") ).getText();

					pivot.print_pivot(pivot1Column_,pivot2Column_,inExcel,outputDir);
					
					
				} catch (Exception e2) {
					// TODO: handle exception
				}
				
				pivotButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
			});
			pivotButton.setOnMouseEntered(e->{
				pivotButton.setStyle("-fx-background-color:#e6ccff; -fx-border-color:#52527a;");
			});
			pivotButton.setOnMouseExited(e ->{
				pivotButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
			});
			//생성버튼 END
			primaryStage.setTitle("사진대지 생성");
			primaryStage.initStyle(StageStyle.UTILITY);
			primaryStage.setScene(scene);
			primaryStage.show();
			
			 
		} catch(Exception e) {
			StringWriter sw = new StringWriter();
			e.printStackTrace(new PrintWriter(sw));
			String exceptionAsString = sw.toString();

			ExceptionCheck exx = new ExceptionCheck();
			try {
				exx.ExceptionCall(exceptionAsString);
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
		}
	}
	

	private void checkPictureFileIsExists(List<DmgStateAndPicture> dmgStateAndPictures){
		if( inPictureDir == null ) return;

		HashMap<String, File> fullFileNames = new HashMap<String,File>();
		File[] filesInPictureDir = inPictureDir.listFiles();

		if(filesInPictureDir != null){
			for(File pictureFile :  filesInPictureDir){
				String fullFileName = pictureFile.getName();
				int lastDot = fullFileName.lastIndexOf('.');
				String nameWithoutExtension = fullFileName.substring(0,lastDot);
				fullFileNames.put(nameWithoutExtension, pictureFile);
			}
		}

		for(DmgStateAndPicture dmgStatPic : dmgStateAndPictures){
			File picFile = fullFileNames.get(dmgStatPic.getPictureFileNameInExcel());
			if(  picFile != null  ){
				dmgStatPic.setPictureFile(picFile);
			}
		}
		return;
	}
	
	private HashMap<String, List<DmgStateAndPicture>> getDSPsDuplicatedOnPictureNumber(List<DmgStateAndPicture> dmgStateAndPictures,String sheetnum){
		HashMap<String, List<DmgStateAndPicture>> duplicationObjs = new HashMap<String, List<DmgStateAndPicture>>();
		for(DmgStateAndPicture dsp :dmgStateAndPictures){
			List<DmgStateAndPicture> dspByPFileName = duplicationObjs.get(dsp.getPictureFileNameInExcel());
			if(dspByPFileName == null){
				dsp.setSheetnum(sheetnum);
				List<DmgStateAndPicture> addingElm = new ArrayList<DmgStateAndPicture>();
				addingElm.add(dsp);
				duplicationObjs.put(dsp.getPictureFileNameInExcel(), addingElm);
			}else{
				dsp.setSheetnum(sheetnum);
				List<DmgStateAndPicture> storedElm = duplicationObjs.get(dsp.getPictureFileNameInExcel());
				storedElm.add(dsp);
			}
		}
/*		
		Iterator<Entry<String, List<DmgStateAndPicture>>>  iter = duplicationObjs.entrySet().iterator();
		
		while(iter.hasNext()){
			Entry<String, List<DmgStateAndPicture>> elm = iter.next();
			System.out.println("--"+elm.getKey()+"--");
			List<DmgStateAndPicture> elmVal = elm.getValue();
			for(DmgStateAndPicture dsp : elmVal){
				System.out.print(" position : " + dsp.getPosition());
				System.out.print(", content : " + dsp.getContent());
				System.out.println(", fileName : " + dsp.getPictureFileNameInExcel());
				
			}
			
		}
*/		
			
		return duplicationObjs;
	}
		 
	public static void main(String[] args) {
		launch(args);
	}	
}