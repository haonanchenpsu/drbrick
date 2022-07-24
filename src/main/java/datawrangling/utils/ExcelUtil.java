package datawrangling.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

public class ExcelUtil {

    private ExcelUtil() {
    }

    private static final int TIME = 0;
    private static final int TALK_TURN = 1;
    private static final int SEGMENT = 2;
    private static final int SPEAKER = 3;
    private static final int TEXT = 4;

    /**
     * Generate cells based on the document and output the xlsx file
     *
     * @param paragraphs    the paragraphs in the Word document
     * @param outputFileDir the output file directory
     */
    public static void writeExcel(List<XWPFParagraph> paragraphs, String outputFileDir) {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            // Sheet initialization
            Sheet sheet = initSheet(workbook);

            // Set the default values
            int rowCount = 4; // The table row starts with line 3
            String time = "";
            int talkTurn = 0;
            String speaker = "";
            int segment = 1;
            boolean hasEndingTime = false; // The paragraph has a date in the end
            boolean ignore = false; // The text should be ignored
            boolean isEmo = false; // The text is related to emotion
            List<String> ignoredWords = Util.loadIgnoredWords();

            // Iterate each paragraph
            for (XWPFParagraph paragraph : paragraphs) {

                String text = paragraph.getText().trim(); // Convert the paragraph to string

                // Skip to next paragraph if it is empty
                if (text.isEmpty()) {
                    break;
                }

                Row row = sheet.createRow(rowCount);
                StringBuilder sb = new StringBuilder();
                boolean hasPrevSegment = false;

                // Set the time first if the previous paragraph has an ending date
                if (hasEndingTime) {
                    row.createCell(TIME).setCellValue(time);
                    hasEndingTime = false;
                }

                for (int i = 0; i < text.length(); i++) {
                    // It reaches the end of the paragraph, and it doesn't have ending time or ignored text
                    if (i == text.length() - 1 && (text.charAt(i) != ']' || isEmo || ignore)) {
                        if (!ignore) {
                            sb.append(text.charAt(i));
                        }
                        row.createCell(TEXT).setCellValue(sb.toString().trim());
                        sb.delete(0, sb.length());
                        row.createCell(SPEAKER).setCellValue(speaker);
                        row.createCell(SEGMENT).setCellValue(segment++); // Increase the segment when the paragraph ends
                        row.createCell(TALK_TURN).setCellValue(talkTurn);
                        rowCount++; // Move to next row
                        ignore = false;
                        isEmo = false;
                        break;
                    }

                    switch (text.charAt(i)) {
                        case ':':
                            if (text.charAt(i + 1) == ' ') { // Get the new speaker
                                speaker = sb.toString().trim();
                                sb.delete(0, sb.length());
                                talkTurn++; // Increase the turn talk when there is a new speaker
                                segment = 1;
                            } else {
                                if (!ignore) { // Append if it is for the time format
                                    sb.append(text.charAt(i));
                                }
                            }
                            break;
                        case '[':
                            if (Character.isAlphabetic(text.charAt(i + 1))) { // Check whether the content is emotion
                                if (ExcelUtil.isIgnored(text, i, ignoredWords)) { // Ignore the text if it's not related
                                    ignore = true;
                                } else { // Treat the content in the brackets as normal text
                                    isEmo = true;
                                    sb.append(text.charAt(i));
                                }
                                break;
                            }

                            // Write the segment before the time bracket
                            if (!sb.toString().trim().equals("")) {
                                row.createCell(SPEAKER).setCellValue(speaker);
                                row.createCell(TALK_TURN).setCellValue(talkTurn);
                                row.createCell(SEGMENT).setCellValue(segment++); // Increase the segment when there is a time bracket
                                row.createCell(TEXT).setCellValue(sb.toString().trim());
                                sb.delete(0, sb.length());
                                hasPrevSegment = true;
                            }
                            break;
                        case ']':
                            // Reset ignore value
                            if (ignore) {
                                ignore = false;
                                break;
                            }
                            if (isEmo) {
                                sb.append(text.charAt(i));
                                isEmo = false;
                                break;
                            }
                            // True if there is segment before the bracket
                            if (hasPrevSegment) {
                                row = sheet.createRow(++rowCount); // Move to next row if the content is time
                                hasPrevSegment = false;
                            }
                            time = sb.toString().trim();
                            sb.delete(0, sb.length());
                            // Set the time if it isn't at the end of the paragraph
                            if (i + 1 >= (text.length() - 1)) {
                                hasEndingTime = true;
                            } else {
                                row.createCell(TIME).setCellValue(time);
                            }
                            break;
                        default:
                            // Append the text
                            if (!ignore) {
                                sb.append(text.charAt(i));
                            }
                            break;
                    }
                }
            }

            // Output the file to the destination
            File dest = new File(outputFileDir);
            FileOutputStream outputStream = new FileOutputStream(dest);
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Sheet initialization
     *
     * @param workbook workbook
     * @return sheet after initialization
     */
    private static Sheet initSheet(XSSFWorkbook workbook) {
        Sheet sheet = workbook.createSheet("Sheet1");

        // Set column width
        sheet.setColumnWidth(0, 2500);
        sheet.setColumnWidth(1, 2500);
        sheet.setColumnWidth(2, 2500);
        sheet.setColumnWidth(3, 5000);
        sheet.setColumnWidth(4, 10000);
        sheet.setColumnWidth(5, 5000);
        sheet.setColumnWidth(6, 5000);

        // Create cell style for label (Bold)
        CellStyle labelStyle = workbook.createCellStyle();
        XSSFFont labelFont = workbook.createFont();
        labelFont.setFontName("Arial");
        labelFont.setBold(true);
        labelStyle.setFont(labelFont);

        // Create cell style for centered label (Bold)
        CellStyle centerLabelStyle = workbook.createCellStyle();
        centerLabelStyle.setFont(labelFont);
        centerLabelStyle.setAlignment(HorizontalAlignment.CENTER);

        // Init row 1 for family id
        Row idRow = sheet.createRow(0);
        Cell familyIdLabel = idRow.createCell(0);
        familyIdLabel.setCellValue("Family ID");
        familyIdLabel.setCellStyle(labelStyle);

        // Init row 2 for file name
        Row fileNameRow = sheet.createRow(1);
        Cell fileNameLabel = fileNameRow.createCell(0);
        fileNameLabel.setCellValue("File name");
        fileNameLabel.setCellStyle(labelStyle);

        // Init row 3 for code
        Row codeRow = sheet.createRow(2);
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 5, 12));
        Cell codeLabel = codeRow.createCell(5);
        codeLabel.setCellValue("Code");
        codeLabel.setCellStyle(centerLabelStyle);

        // Init row 4 for labels
        Row headersRow = sheet.createRow(3);

        Cell timeLabel = headersRow.createCell(0);
        timeLabel.setCellValue("Time");
        timeLabel.setCellStyle(labelStyle);

        Cell talkTurnLabel = headersRow.createCell(1);
        talkTurnLabel.setCellValue("Talk turn");
        talkTurnLabel.setCellStyle(labelStyle);

        Cell segmentLabel = headersRow.createCell(2);
        segmentLabel.setCellValue("Segment");
        segmentLabel.setCellStyle(labelStyle);

        Cell speakerLabel = headersRow.createCell(3);
        speakerLabel.setCellValue("Speaker");
        speakerLabel.setCellStyle(labelStyle);

        Cell textLabel = headersRow.createCell(4);
        textLabel.setCellValue("Text");
        textLabel.setCellStyle(labelStyle);

        Cell negativeEmoLabel = headersRow.createCell(5);
        negativeEmoLabel.setCellValue("NegativeEmotion");
        negativeEmoLabel.setCellStyle(labelStyle);

        Cell emoSupLabel = headersRow.createCell(6);
        emoSupLabel.setCellValue("EmotionalSupport");
        emoSupLabel.setCellStyle(labelStyle);

        Cell momLabel = headersRow.createCell(7);
        momLabel.setCellValue("Mom");
        momLabel.setCellStyle(labelStyle);

        Cell dadLabel = headersRow.createCell(8);
        dadLabel.setCellValue("Dad");
        dadLabel.setCellStyle(labelStyle);

        Cell sib1Label = headersRow.createCell(9);
        sib1Label.setCellValue("Sib1");
        sib1Label.setCellStyle(labelStyle);

        Cell sib2Label = headersRow.createCell(10);
        sib2Label.setCellValue("Sib2");
        sib2Label.setCellStyle(labelStyle);

        Cell parnmLabel = headersRow.createCell(11);
        parnmLabel.setCellValue("Par_NM");
        parnmLabel.setCellStyle(labelStyle);

        Cell ynmLabel = headersRow.createCell(12);
        ynmLabel.setCellValue("Y_NM");
        ynmLabel.setCellStyle(labelStyle);
        return sheet;
    }

    public static boolean isIgnored(String paragraph, int startIndex, List<String> ignoredWords) {
        int endIndex = paragraph.length() - 1;
        for (int i = startIndex; i < paragraph.length(); i++) {
            if (paragraph.charAt(i) == ']' || Character.isDigit(paragraph.charAt(i))) {
                endIndex = i;
                break;
            }
        }
        for (String word : ignoredWords) {
            if (word.equalsIgnoreCase(paragraph.substring(startIndex + 1, endIndex).trim())) {
                return true;
            }
        }
        return false;
    }
}
