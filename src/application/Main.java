package application;
 
import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

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
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.ToggleGroup;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.util.Callback;

public class Main extends Application {
	
	static File inExcel;	// 입력 엑셀
	static File inPictureDir;	// 입력 그림 폴더
	static final String initialWorkingDir;	//현재 프로그램의 작업 디렉토리
	static boolean check_img = true; //사진 중복 체크용
	static{
		initialWorkingDir = System.getProperty("user.dir");
	};
	
	@Override
	public void start(Stage primaryStage) {
		try {
			FXMLLoader loader = new FXMLLoader();
			loader.setLocation(getClass().getResource("/resources/Main.fxml"));
			ObservableMap<String, Object> mainFXMLNamespace =  loader.getNamespace();
			Scene scene = new Scene(loader.load());
			
			ExcelReport excel = new ExcelReportXSSF();
			Util util = new Util();
			
			ProgressEventHandler progressEventHandler = createProgressEventHandler();//작업진행 팝업의 이벤트 핸들링 객체 생성

			//엑셀선택 버튼 START
			Button setExcelButton = (Button) mainFXMLNamespace.get("SetInputExcelButton");
			setExcelButton.setOnMouseClicked(e -> {
				setExcelButton.setStyle("-fx-background-color:#e6ccff; -fx-border-color:#52527a;");
				FileChooser fileChooser = new FileChooser();
				inExcel = (fileChooser.showOpenDialog(primaryStage));
				Label excelPathLabel = (Label)mainFXMLNamespace.get("ExcelPathLabel");
				
				//xlsx파일만 선택하도록함.
				String file_root = inExcel.getAbsolutePath();
				String[] directoryName = file_root.split("\\\\"); 
				String fileName = directoryName[directoryName.length -1];
				if(inExcel != null){
					if(!fileName.contains(".xlsx")){
						ExceptionCheck exx = new ExceptionCheck();
						try {
							exx.ExceptionCall("xlsx확장자 파일만 변환이 가능합니다.\n 이외의 형식은 변환해서 넣어주시기 바랍니다.");
							inExcel = null;
							return;
						} catch (Exception e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
					}
				}
				
				excelPathLabel.setText(inExcel.getAbsolutePath());
				setExcelButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
				check_img = true; // 사진 중복체크용
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

				check_img = true; // 사진 중복체크용
				
				List check_pic_num = new ArrayList<>();
				
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
					HashMap<String, List<DmgStateAndPicture>> dupObjs = getDSPsDuplicatedOnPictureNumber(dmgStateAndPictureSheet); 
					
					sheetCol.setCellValueFactory(new PropertyValueFactory<DmgStateAndPicture,String>("sheetnum"));
					positionCol.setCellValueFactory(new PropertyValueFactory<DmgStateAndPicture,String>("position"));
					pictureNoCol.setCellValueFactory(new PropertyValueFactory<DmgStateAndPicture,String>("pictureFileNameInExcel"));
					contentCol.setCellValueFactory(new PropertyValueFactory<DmgStateAndPicture,String>("content"));
					pictureFile.setCellValueFactory(new PropertyValueFactory<DmgStateAndPicture,String>("pictureFile"));
					
					pictureNoCol.setCellFactory(new Callback<TableColumn<String, String>, TableCell<String, String>>() {
			            @Override
			            public TableCell call(TableColumn p) {
			                return new TableCell<String, String>() {
			                    @Override
			                    public void updateItem(final String item, final boolean empty) {
			                        super.updateItem(item, empty);//*don't forget!
			                        if (item != null) {
			                            setText(item);
			                            if (item.startsWith("중복")) {
			                                setStyle("-fx-background-color: red; -fx-text-fill: white;");
			                            }else{
			                            	setStyle("");
			                            }
			                        } else {
			                            setText(null);
			                        }
			                    }
			                };
			            }
			        });

					for(DmgStateAndPicture dmgStatPic  : dmgStateAndPictureSheet){
						dataList.add(dmgStatPic);
						//숫자 중복 체크
						String check_number = dmgStatPic.getPictureFileNameInExcel().toString()+Integer.toString(dmgStatPic.getSheetnum());
						if(check_pic_num.contains(check_number)){
							dmgStatPic.setPictureFileNameInExcel("중복/"+dmgStatPic.getPictureFileNameInExcel().toString());
							check_img = false; // 사진 중복체크용
						}else{
							check_pic_num.add(check_number);
						}
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
				if(check_img){
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
					String selectedPrintType =selectedRB.getUserData().toString(); //xls  또는 xlsx
					String selectedOutputType = "xlsx";
					
					ExcelReport outExcel = null;
					if(selectedOutputType.equals("xls")){
						outExcel = excel;
					}else{
						outExcel = new ExcelReportXSSF();
						outExcel.setDmgStateAndPictures(excel.getDmgStateAndPictures());
					}
					String pivot1Column_ = null;
					String pivot2Column_ = null;
					
					if(outExcel.getDmgStateAndPictures() == null ) {
						//엑셀 컬럼알파벳을 번호로 변환
						String positionColumn_ =  ( (TextField) mainFXMLNamespace.get("PositionColumnTextField") ).getText();
						String contentColumn_ =  ( (TextField) mainFXMLNamespace.get("ContentColumnTextField") ).getText();
						String pictureNoColumn_ =  ( (TextField) mainFXMLNamespace.get("PictureNoColumnTextField") ).getText();
						pivot1Column_ =  ( (TextField) mainFXMLNamespace.get("Pivot1NoColumnTextField") ).getText();
						pivot2Column_ =  ( (TextField) mainFXMLNamespace.get("Pivot2NoColumnTextField") ).getText();
						int positionColNo = util.decodeToDecimal(positionColumn_);
						int pictureNoColNo = util.decodeToDecimal(pictureNoColumn_);
						int contentColNo = util.decodeToDecimal(contentColumn_);
	
						outExcel.readExcel(contentColNo, pictureNoColNo, positionColNo, inExcel);
						
					}else{
						pivot1Column_ =  ( (TextField) mainFXMLNamespace.get("Pivot1NoColumnTextField") ).getText();
						pivot2Column_ =  ( (TextField) mainFXMLNamespace.get("Pivot2NoColumnTextField") ).getText();
					}
					//outExcel.execute(inPictureDir, outputDir,inExcel,pivot1Column_,pivot2Column_,selectedPrintType, progressEventHandler);
					String pictureNoColumn_ =  ( (TextField) mainFXMLNamespace.get("PictureNoColumnTextField") ).getText();
					int pictureNoColNo = util.decodeToDecimal(pictureNoColumn_);
					//자꾸 0이 나와서 변경한부분
					
					//새 스레드에서 작업을 실행하기 위해 변경
					outExcel.setInfoBeforeExecution(inPictureDir, outputDir,inExcel,pivot1Column_,pivot2Column_,pictureNoColNo,selectedPrintType, progressEventHandler);//하이퍼링크때문에 추
					Runnable runnableExcel = (Runnable) outExcel;
					Thread executeThread = new Thread(runnableExcel);
					executeThread.start();
					
	
					executeButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
				}else{
					String exceptionAsString = "충복된 사진명을 수정해주세요";

					ExceptionCheck exx = new ExceptionCheck();
					try {
						exx.ExceptionCall(exceptionAsString);
					} catch (Exception e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
				}
			});
			executeButton.setOnMouseEntered(e->{
				executeButton.setStyle("-fx-background-color:#e6ccff; -fx-border-color:#52527a;");
			});
			executeButton.setOnMouseExited(e ->{
				executeButton.setStyle("-fx-background-color: #b3b3cc; -fx-border-color: #52527a;");
			});
			//생성버튼 END
	
			primaryStage.setTitle("사진대지 생성");
			primaryStage.initStyle(StageStyle.UTILITY);
			primaryStage.setScene(scene);
			primaryStage.show();
			
			//추가팝업창
			Stage popupStage = new Stage();
	        VBox box = new VBox();

		    Scene popup_scene = new Scene(box);
		    
		    Image main_image = new Image(getClass().getResourceAsStream("/image/mainimage.jpg"));
	        ImageView view_image = new ImageView(main_image);
		    
	        box.getChildren().add(0,view_image);
			
	        popupStage.setTitle("문의정보");
	        popupStage.setScene(popup_scene);
	        popupStage.show();	
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
	
	private HashMap<String, List<DmgStateAndPicture>> getDSPsDuplicatedOnPictureNumber(List<DmgStateAndPicture> dmgStateAndPictures){
		HashMap<String, List<DmgStateAndPicture>> duplicationObjs = new HashMap<String, List<DmgStateAndPicture>>();
		for(DmgStateAndPicture dsp :dmgStateAndPictures){
			List<DmgStateAndPicture> dspByPFileName = duplicationObjs.get(dsp.getPictureFileNameInExcel());
			if(dspByPFileName == null){
				List<DmgStateAndPicture> addingElm = new ArrayList<DmgStateAndPicture>();
				addingElm.add(dsp);
				duplicationObjs.put(dsp.getPictureFileNameInExcel(), addingElm);
			}else{
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
	
	
	public ProgressEventHandler createProgressEventHandler() throws IOException{
		Stage primaryStage = new Stage();
		FXMLLoader loader = new FXMLLoader();
		loader.setLocation(getClass().getResource("/resources/progress.fxml"));
		ObservableMap<String, Object> progressFXMLNamespace =  loader.getNamespace();
		Scene scene = new Scene(loader.load());

		ProgressEventHandler progressEventHandler = new ProgressEventHandler(primaryStage, scene, progressFXMLNamespace);
		primaryStage.setScene(scene);
	    primaryStage.setTitle("작업이 진행중입니다.");

		return (ProgressEventHandler) progressEventHandler;
	}

	
	
}