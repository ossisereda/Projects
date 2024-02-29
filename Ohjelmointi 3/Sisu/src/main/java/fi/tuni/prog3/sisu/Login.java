package fi.tuni.prog3.sisu;

import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.event.ActionEvent;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.stage.Stage;

/**
 * Vastaa suurelta osin kirjautumisnäkymän ulkomuodosta ja toiminnasta
 * @author Ossi
 */
public class Login {
    
    /**
     * Luo tiedot tunnistautumisikkunaan, jossa kysytään käyttäjän
     * opiskelijanumeroa
     * @param stage
     * @return VBox
     */
    public static VBox getLoginView(Stage stage) {
        // Luo tiedot tunnistautumisikkunaan, jossa kysytään opiskelijanumeroa
        
        VBox vbox = new VBox();
        vbox.setSpacing(10);
        vbox.setStyle("-fx-background-color: #7B1FA2;");
        vbox.setAlignment(Pos.CENTER);
        
        // Muokkaa näkymän tekstejä lukemisen helpottamiseksi
        Label first = new Label("Tunnistautuminen");
        first.setFont(Font.font("Helvetica", 20));
        first.setTextFill(Color.WHITE);
        Label second = new Label("Syötä opiskelijanumero:");
        second.setTextFill(Color.WHITE);
        var textField = textField();
        
        vbox.getChildren().addAll(first, second, textField, loginButton(stage, textField));
        
        return vbox; 
    }
    
    /**
     * Luo kirjautumispainikkeen, joka opiskelijanumeron perusteella avaa
     * uuden ikkunan, jossa tutkintotietoja voi muokata
     * @param stage
     * @param textField
     * @return Button
     */
    private static Button loginButton(Stage stage, TextField textField) {
        //Luo kirjautumispainikkeen, joka opiskelijanumeron perusteella
        //avaa uuden ikkunan, jossa tutkintotietoja voi muokata.
              
        Button button = new Button("Kirjaudu");       
        button.setOnAction((ActionEvent event) -> {
            // Napin painallus sulkee vanhan stagen ja avaa käyttäjänäkymän
            // ilmoitetun opiskelijanumeron pohjalta

            stage.close();
            
            String textFieldContent = textField.getText();
            Stage newStage = null;
            try {
                newStage = getUIView.UIStage(textFieldContent);
            } catch (IOException ex) {
                Logger.getLogger(Login.class.getName()).log(Level.SEVERE, null, ex);
            }
            newStage.show();
            
        });
        
        return button;
    }
    
    /**
     * Luo fieldin, jossa kysytään käyttäjän opiskelijanumeroa
     * @return TextField
     */  
    private static TextField textField() {
        // Luo fieldin, jossa kysytään opiskelijanumeroa
        
        TextField number = new TextField("XXXXXX");
        number.setMaxWidth(100);
        number.setAlignment(Pos. CENTER);
        
        return number;
    }
}

