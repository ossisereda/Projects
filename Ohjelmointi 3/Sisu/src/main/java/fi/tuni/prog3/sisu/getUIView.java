package fi.tuni.prog3.sisu;

import com.fasterxml.jackson.core.JsonProcessingException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.TreeSet;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.application.Platform;
import javafx.event.ActionEvent;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.stage.Stage;
import javafx.scene.control.CheckBox;

/**
 * Luo käyttäjänäkymän, joka tulee esiin kun kirjautumispainiketta painetaan
 * @author Ossi
 */

public class getUIView {
    private static TreeSet<StudyModule> kurssiJaModuuliLista = new TreeSet<>(); 

    /**
     * Ottaa parametrikseen kirjautumisnäkymässä kysytyn opiskelijanumeron 
     * ja luo sen pohjalta uuden käyttäjänäkymän
     * @param textFieldContent
     * @return
     * @throws IOException 
     */
    public static Stage UIStage(String textFieldContent) throws IOException {
        // Luo käyttäjänäkymän opiskelijanumeron perusteella

        Stage stage = new Stage();
        stage.setTitle("SISU 2.0 Käyttäjänäkymä - " + textFieldContent);

        BorderPane root = new BorderPane();
        root.setPadding(new Insets(10, 10, 10, 10));
        root.setCenter(getCenterHbox());

        var quitButton = getQuitButton();
        BorderPane.setMargin(quitButton, new Insets(10, 10, 0, 10));
        root.setBottom(quitButton);
        BorderPane.setAlignment(quitButton, Pos.TOP_RIGHT);

        Scene scene = new Scene(root, 1270, 600);
        stage.setScene(scene);

        quitButton.setOnAction((ActionEvent event) -> {
            stage.close();
        });

        stage.show();

        return stage;
    }

    /**
     * Luo käyttäjänäkymän vasemman puoliskon
     * @return
     * @throws IOException 
     */
    private static VBox getLeftVBox() throws IOException {
        VBox leftVBox = new VBox();
        leftVBox.setStyle("-fx-background-color: #7B1FA2;");
        leftVBox.setPadding(new Insets(10));
        leftVBox.setSpacing(10);
        Label prog = new Label("Tutkintokokonaisuus");
        prog.setFont(Font.font("Helvetica", 20));
        prog.setTextFill(Color.WHITE);
        leftVBox.getChildren().add(prog);

        // Lukee tutkintorakenteiden tiedot Sisusta ja listaa ne näkymään buttoneina
        TreeSet<Programme> programme = new SisuApi().getDegreeProgrammes();

        for (Programme i : programme) {
            Button progButton = new Button(i.getName() + ": " + i.getMinCredits() + " op");
            leftVBox.getChildren().add(progButton);
            
            progButton.setOnAction((event )->{
            Programme k = null;
            try {
                k = new SisuApi().getModule(i, 0);
            } catch (JsonProcessingException ex) {
                Logger.getLogger(getUIView.class.getName()).log(Level.SEVERE, null, ex);
            }
            
            kurssiJaModuuliLista = k.getStudyModule();  
            });
        }   

        // Luo ikkunan, jossa tietoja voi tarpeen vaatiessa rullata näkyviin
        ScrollPane scrollPane = new ScrollPane();
        scrollPane.setContent(leftVBox);
        scrollPane.setVbarPolicy(ScrollPane.ScrollBarPolicy.AS_NEEDED);
        scrollPane.setHbarPolicy(ScrollPane.ScrollBarPolicy.AS_NEEDED);
        scrollPane.setFitToWidth(false);
        scrollPane.setPrefViewportWidth(580);
        
        return new VBox(scrollPane);
    }

    /**
     * Luo käyttäjänäkymän oikean puoliskon
     * @return 
     */
    private static VBox getRightVBox() {
        VBox rightVBox = new VBox();
        rightVBox.setStyle("-fx-background-color: #7B1FA2;");
        rightVBox.setPadding(new Insets(10));
        rightVBox.setSpacing(10);

        // Luo ylänurkkaan boxin, johon asetetaan tekstilogo
        HBox logoHBox = new HBox();
        logoHBox.setAlignment(Pos.TOP_RIGHT);
        Label uniLabel = new Label("Tampereen yliopisto");
        uniLabel.setFont(Font.font("Helvetica", 30));
        uniLabel.setTextFill(Color.WHITE);
        logoHBox.getChildren().add(uniLabel);

        // Luo boxin, johon kerätään tutkinnon rakenne
        HBox infoHBox = new HBox();
        infoHBox.setAlignment(Pos.CENTER_LEFT);
        infoHBox.setSpacing(10);

        VBox namesVBox = new VBox();
        namesVBox.setAlignment(Pos.TOP_LEFT);
        namesVBox.setSpacing(10);

        // Luo mahdollisuuden selata tietoja alemmas, jos ne eivät mahdu ruutuun
        ScrollPane scrollPane = new ScrollPane();
        scrollPane.setContent(namesVBox);
        scrollPane.setVbarPolicy(ScrollPane.ScrollBarPolicy.AS_NEEDED);
        scrollPane.setFitToWidth(false);
        scrollPane.setPrefSize(550, 400);

        // Luo painikkeen, jota painamalla valitusta 
        // tutkinnosta saadaan tiedot näkyviin
        Button infoButton = new Button("INFO");
        infoButton.setOnAction((ActionEvent e) -> {
            getCourses(namesVBox);
        });

        infoHBox.getChildren().addAll(infoButton, scrollPane);
        rightVBox.getChildren().addAll(logoHBox, infoHBox);

        return rightVBox;
    }    

    /**
     * Yhdistää käyttäjänäkymän vasemman ja oikean puoliskon yhdeksi
     * @return
     * @throws IOException 
     */
    private static HBox getCenterHbox() throws IOException {
        //Yhdistää käyttäjänäkymän vasemman ja oikean puoliskon yhdeksi        
        HBox centerHBox = new HBox(20);
        centerHBox.getChildren().addAll(getLeftVBox(), getRightVBox());

        return centerHBox;
    }

    /**
     * Luo napin, jota painamalla ohjelma tulee ensisijaisesti lopettaa
     * @return 
     */
    private static Button getQuitButton() {
        // Luo napin, jota painamalla ohjelma saadaan lopetettua        
        Button button = new Button("Tallenna ja lopeta");
        button.setOnAction((ActionEvent event) -> {
            Platform.exit();
        });      

        return button;
    }

    /**
     * Luo käyttöliittymään rakenteen, jossa kurssit on jaettu moduuleittain
     * @param namesVBox 
     */
    private static void getCourses(VBox namesVBox) {
        // Luo rakenteen, jossa kurssit on jaoteltu moduuleittain

        namesVBox.getChildren().clear();

        for (StudyModule studymodule : kurssiJaModuuliLista) {
            Button moduleButton = new Button(studymodule.getName() + ": väh. " + studymodule.getMinCredits() + " op");
            if (studymodule.isRequirementFulfilled()) {
                moduleButton.setStyle("-fx-background-color: #00ff00;");
            }else{
                moduleButton.setStyle(null);
            }

            HBox moduleBox = new HBox(moduleButton);

            final List<Node> studymoduleDetails = new ArrayList<>();
            boolean[] showDetails = { false };

            for (CourseUnit course : studymodule.getCourses()) {
                CheckBox courseCheckbox = new CheckBox();
                Label nameLabelCourses = new Label("  " + course.getName() + ": " + course.getMinCredits() + " op");
                HBox courseBox = new HBox(courseCheckbox, nameLabelCourses);
                studymoduleDetails.add(courseBox);

                if (course.isChecked()){
                    courseCheckbox.setSelected(true);
                }

                courseCheckbox.setOnAction((e) -> {
                    course.changeChecked(courseCheckbox.isSelected());
                    moduleButton.setText(studymodule.getName() + ": väh. " + studymodule.getMinCredits() + 
                        " op / suoritettu " + studymodule.getCoursesPoints() + " op");

                    // Muuttaa moduulinapin vihreäksi, jos vähimmäisopintopisteet
                    // täyttyvät moduulikohtaisista kurssivalinnoista
                    if (studymodule.isRequirementFulfilled()) {
                        moduleButton.setStyle("-fx-background-color: #00ff00;");
                    }else{
                        moduleButton.setStyle(null);
                    }
                });
            }   

            // 
            moduleButton.setOnAction((ActionEvent ev) -> {
                if (showDetails[0]) {
                    namesVBox.getChildren().removeAll(studymoduleDetails);
                    showDetails[0] = false;
                } else {
                    int index = namesVBox.getChildren().indexOf(moduleBox);
                    namesVBox.getChildren().addAll(index + 1, studymoduleDetails);
                    showDetails[0] = true;
                }
            });

            namesVBox.getChildren().add(moduleBox);
        }
    }
}