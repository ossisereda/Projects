package fi.tuni.prog3.sisu;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.layout.BorderPane;
import javafx.stage.Stage;

/**
 * Käynnistää koko ohjelman
 * @author Ossi
 */

public class Sisu extends Application {

    /**
     * Luo edellytykset ohjelman aloittamiselle
     * @param stage 
     */ 
    @Override
    public void start(Stage stage) {

        BorderPane root = new BorderPane();
        root.setPadding(new Insets(10, 10, 10, 10));
        root.setCenter(Login.getLoginView(stage));
        Scene scene = new Scene(root, 400, 200); 
        stage.setTitle("SISU 2.0 Kirjautuminen");
        stage.setScene(scene);
        stage.show();
    }

    /**
     * Aloittaa ohjelman
     * @param args 
     */
    public static void main(String[] args) {
     
        launch();
    }
}