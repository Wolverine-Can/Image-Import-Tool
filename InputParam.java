import javafx.scene.control.TextArea;

public class InputParam {
    private final String workbookPath;
    private final String imageFolderPath;
    private final String imagePrefix;
    private final String imageExtension;
    private final String imageWidth;
    private final String imageHeight;

    public InputParam(String workbookPath, String imageFolderPath, String imagePrefix, String imageExtension,
                      String imageWidth, String imageHeight) {
        this.workbookPath = workbookPath;
        this.imageFolderPath = imageFolderPath;
        this.imagePrefix = imagePrefix;
        this.imageExtension = imageExtension;
        this.imageWidth = imageWidth;
        this.imageHeight = imageHeight;
    }

    public String getWorkbookPath() {
        return workbookPath;
    }

    public String getImageFolderPath() {
        return imageFolderPath;
    }

    public String getImagePrefix() {
        return imagePrefix;
    }

    public String getImageExtension() {
        return imageExtension;
    }

    public String getImageWidth() {
        return imageWidth;
    }

    public String getImageHeight() {
        return imageHeight;
    }

}
