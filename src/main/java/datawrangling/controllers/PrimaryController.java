package datawrangling.controllers;

import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import datawrangling.utils.ExcelUtil;
import datawrangling.utils.Util;

import java.util.List;

public class PrimaryController {
    @FXML
    private TextField inputPath;
    @FXML
    private TextField outputPath;
    @FXML
    private Label msgLabel;

    /**
     * Transcribe the Word documents
     */
    @FXML
    private void transcribe() {
        msgLabel.setText("");

        // Get the input-dir and output-dir from GUI
        String inputFolderDir = Util.convertFilePath(inputPath.getText().trim());
        String outputFolderDir = Util.convertFilePath(outputPath.getText().trim());

        // Notify user when the input-dir is empty
        if (inputFolderDir.isEmpty()) {
            msgLabel.setText("The input folder directory is empty. Please try again.");
            return;
        }

        // Notify user when the input-dir does not exist
        if (!Util.sanityCheckFolderDir(inputFolderDir, true, false)) {
            msgLabel.setText("The input folder directory does not exist. Please try again.");
            return;
        }

        // Auto-generate output folder directory if it's empty
        if (outputFolderDir.isEmpty()) {
            String parentFile = inputFolderDir.substring(0, inputFolderDir.substring(0, inputFolderDir.length() - 1).lastIndexOf('/'));
            String fileName = inputFolderDir.substring(parentFile.length(), inputFolderDir.lastIndexOf('/'));
            outputFolderDir = parentFile.concat(fileName + " Result" + "/");
            Util.sanityCheckFolderDir(outputFolderDir, false, true);
        } else {
            if (!Util.sanityCheckFolderDir(outputFolderDir, false, false)) {
                msgLabel.setText("The output folder directory is incorrect. (You can leave it blank if the issue can't be solved.)");
                return;
            }
        }

        // Convert docx files into xlsx files
        List<String> inputDocNames = Util.getFileNames(inputFolderDir);
        for (String fileName : inputDocNames) {
            List<XWPFParagraph> paragraphs = Util.readDocx(inputFolderDir + fileName);
            ExcelUtil.writeExcel(paragraphs, outputFolderDir + fileName.substring(0, fileName.lastIndexOf('.')) + ".xlsx");
        }
        msgLabel.setText("Finish converting.");
    }
    @FXML
    private void exit() {
        System.exit(0);
    }
}