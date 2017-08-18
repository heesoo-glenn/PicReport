package application;

import javafx.collections.ObservableMap;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextArea;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;
import javafx.stage.Stage;

public class ExceptionCheck {

	 public void ExceptionCall(String e) throws Exception{

		Stage primaryStage = new Stage();
        VBox box = new VBox();

	    Scene scene = new Scene(box, 900, 500);
	    
	    Image main_image = new Image(getClass().getResourceAsStream("/image/mainimage.jpg"));
        ImageView view_image = new ImageView(main_image);
	    
        box.getChildren().add(0,view_image);
	 
        Text text_main = new Text("\n아래의 메시지를 복사해서 이메일로 첨부해주세요\n");
	    text_main.setStyle("-fx-font-size: 20;");
        
	    box.getChildren().add(1,text_main);
	    	    
	    
	    ScrollPane root = new ScrollPane();
	    
	    TextArea textArea = new TextArea();
	    box.getChildren().add(2, textArea );
        textArea.setStyle("-fx-font-size: 15;");
        textArea.setText(e);
        textArea.deselect();
	    primaryStage.setTitle("Error");
        primaryStage.setScene(scene);
        primaryStage.show();	    
	 }
}
