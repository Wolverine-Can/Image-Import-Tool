import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.text.Font;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;

public class Main extends Application {
    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage stage) throws Exception {
        stage.setTitle("Image Import Tool");
        stage.setResizable(false);
        stage.getIcons().add(new Image(getClass().getResourceAsStream("icon.png")));

        /* Components */
        //file chooser
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Select Workbook");
        fileChooser.getExtensionFilters().addAll(new FileChooser.ExtensionFilter("Workbook", "*.xls", "*.xlsx"));
        Label fileChooserLabel = new Label("No workbook selected");
        fileChooserLabel.setMaxWidth(200);
        Button fileChooserButton = new Button("Select Workbook");
        fileChooserButton.setPrefWidth(135);
        fileChooserButton.setOnAction(e -> {
            File file = fileChooser.showOpenDialog(stage);
            if (file != null) {
                fileChooserLabel.setText(file.getAbsolutePath());
            }
        });
        HBox fileChooserHBox = new HBox(10);
        fileChooserHBox.getChildren().addAll(fileChooserButton, fileChooserLabel);

        //directory chooser
        DirectoryChooser directoryChooser = new DirectoryChooser();
        Label directoryChooserLabel = new Label("No directory selected");
        Button directoryChooserButton = new Button("Select Image Directory");
        directoryChooserButton.setOnAction(e -> {
            File file = directoryChooser.showDialog(stage);
            if (file != null) {
                directoryChooserLabel.setText(file.getAbsolutePath());
            }
        });
        HBox directoryChooserHBox = new HBox(10);
        directoryChooserHBox.getChildren().addAll(directoryChooserButton, directoryChooserLabel);

        //prefix input
        Label prefixLabel = new Label("Image Name Prefix:         ");
        TextField prefixInput = new TextField();
        prefixInput.setPromptText("prefix");
        HBox prefixHBox = new HBox(10);
        prefixHBox.getChildren().addAll(prefixLabel, prefixInput);

        //extension input
        Label extensionLabel = new Label("Image Name Extension:   ");
        TextField extensionInput = new TextField();
        extensionInput.setPromptText(".jpg or .png");

        HBox extensionHBox = new HBox(10);
        extensionHBox.getChildren().addAll(extensionLabel, extensionInput);

        //image size input
        Label imageSizeLabel = new Label("Set Image Size:                ");
        TextField imageColInput = new TextField();
        imageColInput.setPrefWidth(62);
        imageColInput.setPromptText("Columns");
        Label timesSign = new Label("x");
        TextField imageRowInput = new TextField();
        imageRowInput.setPrefWidth(62);
        imageRowInput.setPromptText("Rows");
        HBox imageSizeHBox = new HBox(10);
        imageSizeHBox.getChildren().addAll(imageSizeLabel, imageColInput, timesSign, imageRowInput);

        //message area
        TextArea messageArea = new TextArea();
        messageArea.setEditable(false);
        messageArea.setWrapText(true);
        messageArea.setMaxHeight(Double.MAX_VALUE);
        messageArea.appendText("Welcome to Image Import Tool! You can use this tool to easily import multiple images to your workbook.\n" +
                "Instruction:\n" +
                "1. Select workbook and the directory of images.\n" +
                "2. Fill in image prefix and extension (for example, " + '"' + "scope_23.png" + '"' + " has a prefix of " + '"' + "scope_" + '"' + " and extension of " + '"' + ".png" + '"' +
                ". Meanwhile, in your workbook, put " + '"' + "`23" + '"' + " (note a backtick is needed) " + "in the cell where the upper left conner of the image will locate. " +
                "This tool supports .jpg and .png formats.\n" +
                "3. Fill in image size.\n" +
                "4. Click Import Images button to start importing.\n" +
                "Developer: Can Deng\n");

        //import button
        Button importButton = new Button();
        importButton.setId("importButton");
        importButton.setPrefWidth(150);
        importButton.setPrefHeight(80);
        importButton.setFont(new Font(18));
        importButton.setText("Import Images");
        importButton.setOnAction(e -> {
            String workbookPath = fileChooserLabel.getText();
            String imageFolderPath = directoryChooserLabel.getText();
            String imagePrefix = prefixInput.getText();
            String imageExtension = extensionInput.getText();
            String imageWidth = imageColInput.getText();
            String imageHeight = imageRowInput.getText();
            InputParam inputParam = new InputParam(workbookPath, imageFolderPath, imagePrefix,
                                                    imageExtension, imageWidth, imageHeight);
            ImportButtonHandler importButtonHandler = new ImportButtonHandler(inputParam, messageArea);
            importButtonHandler.handle(e);
        });

        VBox upperLeft = new VBox(20);
        upperLeft.setMaxWidth(400);
        upperLeft.setPadding(new Insets(20));
        upperLeft.setAlignment(Pos.CENTER);
        upperLeft.getChildren().addAll(fileChooserHBox, directoryChooserHBox, prefixHBox, extensionHBox, imageSizeHBox);

        VBox upperRight = new VBox(20);
        upperRight.setPadding(new Insets(20));
        upperRight.setAlignment(Pos.BOTTOM_LEFT);
        upperRight.getChildren().addAll(importButton);

        HBox upper = new HBox(20);
        upperRight.setPadding(new Insets(20));
        upperRight.setAlignment(Pos.BOTTOM_CENTER);
        upper.getChildren().addAll(upperLeft, upperRight);

        VBox layout = new VBox(20);
        layout.setPadding(new Insets(20));
        layout.setAlignment(Pos.CENTER);
        layout.getChildren().addAll(upper, messageArea);

        Scene scene = new Scene(layout, 640, 480);
        scene.getStylesheets().add("styles.css");
        stage.setScene(scene);
        stage.show();
    }
}
