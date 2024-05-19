package PdfToExcel.Project2;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class ExcelWriter {

    public static void writeToExcel(List<ContentItem> contentItemList, String filePath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Content Items");

            int rowNum = 0;
            for (ContentItem contentItem : contentItemList) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(contentItem.getTitle());
                row.createCell(1).setCellValue(contentItem.getBody());
            }

            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
                System.out.println("Excel file written successfully.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        // Example list of ContentItem objects
        ContentItem item1 = new ContentItem("Title 1", "Body 1");
        ContentItem item2 = new ContentItem("Title 2", "Body 2");
        ContentItem item3 = new ContentItem("Title 3", "Body 3");
        
        List<ContentItem> contentItemList = Arrays.asList(item1, item2, item3);

        // Path to the output Excel file
        String excelFilePath = "output.xlsx";

        // Write the list of ContentItem objects to Excel
        writeToExcel(contentItemList, excelFilePath);
    }
}

class ContentItem {
    private String title;
    private String body;

    public ContentItem(String title, String body) {
        this.title = title;
        this.body = body;
    }

    public String getTitle() {
        return title;
    }

    public String getBody() {
        return body;
    }
}
