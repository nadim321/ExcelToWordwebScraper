import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
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

public class Militaryaerospace {

    public static void main(String[] args) {
        // Configuration
        String excelPath = "input_urls.xlsx"; // Excel with URLs
        String wordOutputPath = "batch_1_militaryaerospace.docx";

        try {
            // 1. Read URLs from Excel
            List<String> urls = readUrlsFromExcel(excelPath);
            List<String> headlines = readHeadlineFromExcel(excelPath);
            List<String> dates = readDateFromExcel(excelPath);
            List<String> others = readOthersFromExcel(excelPath);

            // 2. Create Word document
            XWPFDocument wordDoc = new XWPFDocument();

            // 3. Process each URL
            for (int i = 0 ; i < urls.size() ; i++) {
                String url = urls.get(i);
                String headLine = headlines.get(i);
                String date = dates.get(i);
                String other = others.get(i);
                System.out.println("Processing: " + url);

                // Scrape content
                String content = scrapeArticle(url);

                // Add to Word doc (new page per URL)
                addToWordDoc(i+1, wordDoc, url,headLine, date,other, content);
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
                Cell cell = row.getCell(7);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    urls.add(cell.getStringCellValue());
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return urls;
    }

    private static List<String> readOthersFromExcel(String filePath) throws IOException {
        List<String> others = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                String headline = "";
                Cell cell = row.getCell(1);
                Cell cell2 = row.getCell(2);
                Cell cell3 = row.getCell(3);
                Cell cell4 = row.getCell(4);
                Cell cell5 = row.getCell(5);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    headline += cell.getStringCellValue() + " - ";
                }
                if (cell2 != null && cell2.getCellType() == CellType.STRING) {
                    headline += cell2.getStringCellValue() + " - ";
                }
                if (cell3 != null && cell3.getCellType() == CellType.STRING) {
                    headline += cell3.getStringCellValue() + " - ";
                }
                if (cell4 != null && cell4.getCellType() == CellType.STRING) {
                    headline += cell4.getStringCellValue() + " - ";
                }
                if (cell5 != null && cell5.getCellType() == CellType.STRING) {
                    headline += cell5.getStringCellValue();
                }
                others.add(headline);
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return others;
    }

    private static List<String> readHeadlineFromExcel(String filePath) throws IOException {
        List<String> headlines = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                Cell cell = row.getCell(6);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    headlines.add(cell.getStringCellValue());
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return headlines;
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
//    private static String scrapeArticle(String url) throws IOException {
//        StringBuilder content = new StringBuilder();
//        try {
//          Document doc = Jsoup.connect(url)
//                  .userAgent("Mozilla/5.0")
//                  .timeout(20000)
//                  .get();
//
//          Elements paragraphs = doc.select("p");
//          boolean isStart = false;
//          for (Element p : paragraphs) {
////              if(p.text().contains(" Comment") && !isStart){
////                  isStart = true;
////                  continue;
////              }
////              if(!isStart){
////                  continue;
////              }
//              if(p.text().contains("Home » ")) {
//                  continue;
//              }
//              if(p.text().contains("Related Posts") || p.text().contains("Report an Issue")){
//                  break;
//              }
//              String text = p.text().trim();
//              if (!text.isEmpty()) {
//                  content.append(text).append("\n\n");
//              }
//          }
//
//      }catch (Exception e){
//          e.printStackTrace();
//      }
//      return content.toString();
//    }


    private static String scrapeArticle(String url) {
        StringBuilder content = new StringBuilder();
        try {
            // Configure connection with proper timeouts and headers
            Document doc = Jsoup.connect(url)
                    .userAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
                    .timeout(20000) // 30 seconds timeout
                    .ignoreHttpErrors(true)
                    .followRedirects(true)
                    .maxBodySize(0) // No limit on page size
                    .execute()
                    .parse();

            // Extract content
            Elements paragraphs = doc.select("p");
            for (Element p : paragraphs) {
                String text = p.text().trim();
                if (!text.isEmpty()) {
                    if(text.contains("Related:") || text.contains("Senior Editor")){
                        continue;
                    }
                    if(text.contains("Editor-in-Chief")){
                        break;
                    }
                    content.append(text).append("\n\n");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return content.toString();
    }

    // Add content to Word doc
    private static void addToWordDoc(int i , XWPFDocument doc, String url, String headline, String date, String other, String content) {
       try {
           // Add URL as heading
           XWPFParagraph urlPara = doc.createParagraph();
           XWPFRun urlRun = urlPara.createRun();
           urlRun.setBold(true);
           urlRun.setFontSize(12);
           urlRun.setText(i+" . Date: " + date);
           urlRun.addCarriageReturn();
           urlRun.setText(other);
           urlRun.addCarriageReturn();
           urlRun.addCarriageReturn();
           urlRun.setText(headline);
           urlRun.addCarriageReturn();
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