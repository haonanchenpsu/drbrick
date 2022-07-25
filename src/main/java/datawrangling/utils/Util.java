package datawrangling.utils;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class Util {
    private static final String IGNORED_WORDS_FILE_DIR = "src/main/resources/IgnoredWords.txt";

    {
        File file = new File(IGNORED_WORDS_FILE_DIR);
        if (!file.exists()) {
            System.out.println("Fail to load word file.");
        }
    }

    private Util() {

    }

    /**
     * Get the list of file names under the input file
     *
     * @return a list of file names
     */
    public static List<String> getFileNames(String inputFolder) {
        File[] inputDocNames = new File(inputFolder).listFiles();

        List<String> fileNames = new ArrayList<>();
        for (File document : inputDocNames) {
            String fileExtension = document.getName().substring(document.getName().lastIndexOf('.'));
            if (fileExtension.equals(".docx")) {
                fileNames.add(document.getName());
            }
        }
        return fileNames;
    }

    public static String convertFilePath(String rawPath) {
        if (rawPath.isEmpty()) {
            return "";
        }
        String path = rawPath.replaceAll("\\\\", "/");
        if (path.charAt(path.length() - 1) != '/') {
            path = path + '/';
        }
        return path;
    }

    /**
     * Get the paragraphs from the document
     *
     * @param fileDir name of the file
     * @return a list of paragraph
     */
    public static List<XWPFParagraph> readDocx(String fileDir) {
        try (FileInputStream fis = new FileInputStream(fileDir)) {
            return new XWPFDocument(fis).getParagraphs();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return Collections.emptyList();
    }

    public static List<String> loadIgnoredWords() {
        File file = new File(IGNORED_WORDS_FILE_DIR);

        List<String> ignoredWords = new ArrayList<>();

        try (BufferedReader reader = new BufferedReader(new FileReader(file))) {
            String word;
            while ((word = reader.readLine()) != null) {
                ignoredWords.add(word.trim());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return ignoredWords;
    }

    public static boolean sanityCheckFolderDir(String dir, boolean isInput, boolean isEmpty) {
        if (isInput) {
            return new File(dir).exists();
        } else {
            if (isEmpty) {
                File outputFolder = new File(dir);
                if (!outputFolder.exists()) {
                    outputFolder.mkdir();
                }
                return true;
            } else {
                return new File(dir).exists();
            }
        }
    }
}
