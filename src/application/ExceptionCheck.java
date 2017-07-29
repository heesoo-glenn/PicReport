package application;

import javafx.collections.ObservableMap;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.stage.Stage;

public class ExceptionCheck {

	 public void ExceptionCall(String e) throws Exception{

		Stage primaryStage = new Stage();
		FXMLLoader loader = new FXMLLoader();
		loader.setLocation(getClass().getResource("/resources/Exception.fxml"));
		ObservableMap<String, Object> mainFXMLNamespace =  loader.getNamespace();
		Scene scene = new Scene(loader.load());
	    
		//¿¡·¯Ã¢
		Label setErrLabel = (Label) mainFXMLNamespace.get("main_text");
				
		setErrLabel.setText(e);
		
		primaryStage.setScene(scene);
	    primaryStage.setTitle("Error");
	    primaryStage.show();	    
	    
	    
	 }
}
