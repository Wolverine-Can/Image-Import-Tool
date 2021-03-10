import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.scene.control.TextArea;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.security.InvalidAlgorithmParameterException;


class ImportButtonHandler implements EventHandler<ActionEvent> {
    private InputParam inputParam;
    private TextArea messageArea;

    public ImportButtonHandler(InputParam inputParam, TextArea messageArea) {
        this.inputParam = inputParam;
        this.messageArea = messageArea;
    }

    public void handle(ActionEvent event) {
        String lineBreaker = "******************************************************************************************************************\n";
        messageArea.appendText(lineBreaker);
        try {
            String workbookPath = inputParam.getWorkbookPath();
            String imageFolderPath = inputParam.getImageFolderPath();
            String imagePrefix = inputParam.getImagePrefix();
            String imageExtension = inputParam.getImageExtension();
            if(inputParam.getImageWidth().equals("") || inputParam.getImageHeight().equals("")) {
                throw new InvalidAlgorithmParameterException();
            }
            int imageWidth = Integer.parseInt(inputParam.getImageWidth());
            int imageHeight = Integer.parseInt(inputParam.getImageHeight());
            FileInputStream fileInputStream = new FileInputStream(new File(workbookPath));
            Workbook workbook = workbookPath.charAt(workbookPath.length() - 1) == 'x' ? new XSSFWorkbook(fileInputStream) : new HSSFWorkbook(fileInputStream);
            fileInputStream.close();

            //Import images
            messageArea.appendText("Started importing\n");
            importImages(workbook, imageFolderPath, imagePrefix, imageExtension, imageWidth, imageHeight, messageArea);

            //Write the Excel file
            File file = new File(workbookPath);
            String newFilePath = file.getParentFile().getPath() + "\\" + "img_imported_" + file.getName();
            messageArea.appendText("Saving to file: " + newFilePath + "\n");
            FileOutputStream fileOut = new FileOutputStream(newFilePath);
            workbook.write(fileOut);
            fileOut.close();
            messageArea.appendText("\nFinished importing\n");
        } catch (InvalidAlgorithmParameterException iapex) {
            messageArea.appendText("Error: Please select file/directory and fill in required fields\n");
        } catch (IOException ioex) {
            messageArea.appendText("Error: Fail to save file. Workbook is opened by other application\n");
        }
    }

    private void importImages(Workbook workbook, String imageFolderPath, String prefix, String extension,
                              int imageWidth, int imageHeight, TextArea messageArea) {
        for(Sheet sheet : workbook) {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if(cell.getCellType() != CellType.STRING) {
                        continue;
                    }
                    try{
                        String cellStr = cell.getStringCellValue();
                        int cellRow = cell.getRowIndex();
                        int cellCol = cell.getColumnIndex();
                        if(cellStr.charAt(0) != '`') {
                            continue;
                        }
                        String fileNum = cellStr.substring(1);
                        String imagePath = imageFolderPath + "\\" + prefix + fileNum + extension;
                        messageArea.appendText("Importing " +  imagePath);
                        File file = new File(imagePath);
                        if(!file.exists()) {
                            throw new IOException(imagePath);
                        }
                        InputStream inputStream = new FileInputStream(imagePath);
                        byte[] bytes = IOUtils.toByteArray(inputStream);
                        int pictureIdx = 0;
                        if(extension.equals(".jpg")) {
                            pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
                        }
                        if(extension.equals(".png")) {
                            pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                        }
                        inputStream.close();
                        CreationHelper helper = workbook.getCreationHelper();
                        Drawing drawing = sheet.createDrawingPatriarch();
                        ClientAnchor anchor = helper.createClientAnchor();
                        anchor.setCol1(cellCol);
                        anchor.setRow1(cellRow);
                        anchor.setCol2(anchor.getCol1() + imageWidth);
                        anchor.setRow2(anchor.getRow1() + imageHeight);
                        Picture pict = drawing.createPicture(anchor, pictureIdx);
                        cell.setBlank();
                        messageArea.appendText("              success!\n");
                    }
                    catch(IOException ioex) {
                        messageArea.appendText("              image not found\n");
                    }
                }
            }
        }
    }
}
