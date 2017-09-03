package com.fetch;

import com.fetch.*;
import javafx.application.Application;
import javafx.fxml.*;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.paint.Color;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

public class MyProgram extends Application{
	
	@Override
    public void start(Stage stage) throws Exception {
        Parent root = FXMLLoader.load(getClass().getResource("excelProgram.fxml"));
        
        Scene scene = new Scene(root,Color.TRANSPARENT);
//        stage.initStyle(StageStyle.UTILITY);
        stage.setScene(scene);
      
        stage.show();
    }
	
	public static void main(String[] args){
		launch(args);		
	}
}
