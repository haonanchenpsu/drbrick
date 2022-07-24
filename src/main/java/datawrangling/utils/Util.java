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
     * @return a list of file names
     */
    public static List<String> getFileNames(String inputFolder) {

        File inputFolderDir = new File(convertFilePath(inputFolder));

        if (!inputFolderDir.exists()) {
            System.out.println("The input folder does not exist.");
            return new ArrayList<>();
        }

        File[] inputDocNames = new File(inputFolder).listFiles();

        List<String> fileNames = new ArrayList<>();
        if (inputDocNames != null && inputDocNames.length > 0) {
            for (File document : inputDocNames) {
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

    public static void addIgnoredWord(String word) {
        File file = new File(IGNORED_WORDS_FILE_DIR);

        try (BufferedWriter writer = new BufferedWriter(new FileWriter(file, true))) {
            writer.write(word + "\r\n");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    private static void updateWordFile(List<String> updatedWordList) {
        File file = new File(IGNORED_WORDS_FILE_DIR);

        try (BufferedWriter writer = new BufferedWriter(new FileWriter(file, true))) {
            clearFile();
            for (String ignoredWord : updatedWordList) {
                writer.write(ignoredWord + "\r\n");
            }
            writer.flush();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    public static void removeIgnoredWord(String word) {
        List<String> ignoredWords = Util.loadIgnoredWords();
        ignoredWords.remove(word);
        updateWordFile(ignoredWords);
    }
    
    private static void clearFile() {
        try (FileWriter fileWriter = new FileWriter(IGNORED_WORDS_FILE_DIR)) {
            fileWriter.write("");
        } catch (Exception e) {
            e.printStackTrace();
        }
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
