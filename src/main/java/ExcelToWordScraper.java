import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelToWordScraper {

    public static void main(String[] args) {
        // Configuration
        String excelPath = "input_urls.xlsx"; // Excel with URLs
        String wordOutputPath = "output_articles.docx";

        try {
            // 1. Read URLs from Excel
            List<String> urls = readUrlsFromExcel(excelPath);
            List<String> headlines = readHeadlineFromExcel(excelPath);
            List<String> dates = readDateFromExcel(excelPath);

            // 2. Create Word document
            XWPFDocument wordDoc = new XWPFDocument();

            // 3. Process each URL
            for (int i = 0 ; i < urls.size() ; i++) {
                String url = urls.get(i);
                String headLine = headlines.get(i);
                String date = dates.get(i);
                System.out.println("Processing: " + url);

                // Scrape content
                String content = scrapeArticle(url);

                // Add to Word doc (new page per URL)
                addToWordDoc(wordDoc, url,headLine, date, content);
                wordDoc.createParagraph().setPageBreak(true); // New page
            }

            // 4. Save Word file
            try (FileOutputStream out = new FileOutputStream(wordOutputPath)) {
                wordDoc.write(out);
            }
            System.out.println("Saved all articles to: " + wordOutputPath);

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
        }
    }

    // Read URLs from Excel (first column)
    private static List<String> readUrlsFromExcel(String filePath) throws IOException {
        List<String> urls = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                Cell cell = row.getCell(3);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    urls.add(cell.getStringCellValue());
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return urls;
    }

    private static List<String> readHeadlineFromExcel(String filePath) throws IOException {
        List<String> urls = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                Cell cell = row.getCell(2);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    urls.add(cell.getStringCellValue());
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return urls;
    }

    private static List<String> readDateFromExcel(String filePath) throws IOException {
        List<String> urls = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                Cell cell = row.getCell(0);
                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    Date date = cell.getDateCellValue();

                    // Format the date to "dd-MM-yyyy"
                    SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
                    String formattedDate = dateFormat.format(date);
                    urls.add(formattedDate);
                }
            }
        }
        return urls;
    }

    // Scrape article content
    private static String scrapeArticle(String url) throws IOException {
        StringBuilder content = new StringBuilder();
        try {
          Document doc = Jsoup.connect(url)
                  .userAgent("Mozilla/5.0")
                  .timeout(10000)
                  .get();

          Elements paragraphs = doc.select("p");
          for (Element p : paragraphs) {
              String text = p.text().trim();
              if (!text.isEmpty()) {
                  content.append(text).append("\n\n");
              }
          }

      }catch (Exception e){
          e.printStackTrace();
      }
      return content.toString();
    }

    // Add content to Word doc
    private static void addToWordDoc(XWPFDocument doc, String url, String headline, String date, String content) {
       try {
           // Add URL as heading
           XWPFParagraph urlPara = doc.createParagraph();
           XWPFRun urlRun = urlPara.createRun();
           urlRun.setBold(true);
           urlRun.setFontSize(12);
           urlRun.setText("Headline: " + headline);
           urlRun.addCarriageReturn();
           urlRun.setText("Date: " + date);
           urlRun.addCarriageReturn();
           urlRun.setText("URL: " + url);
           urlRun.addBreak(); // Optional line break after URL

           // Split content by double newlines to preserve paragraphs
           String[] paragraphs = content.split("\n\n");

           // Add each paragraph to Word doc
           for (String paragraph : paragraphs) {
               XWPFParagraph wordPara = doc.createParagraph();
               XWPFRun run = wordPara.createRun();
               run.setText(paragraph.trim());
               run.addCarriageReturn(); // Add single line break after each paragraph (optional)
           }
       }catch (Exception e){
           e.printStackTrace();
       }
    }
}