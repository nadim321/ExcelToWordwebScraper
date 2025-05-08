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
import java.util.*;

public class Uasvision {

    public static void main(String[] args) {
        // Configuration
        String excelPath = "input_urls.xlsx";
        String wordOutputPath2023 = "batch_1_uasvision_2023.docx";
        String wordOutputPath2024 = "batch_1_uasvision_2024.docx";
        String wordOutputPath2025 = "batch_1_uasvision_2025.docx";

        try {
            List<String> urls = readUrlsFromExcel(excelPath);
            List<String> headlines = readHeadlineFromExcel(excelPath);
            List<String> dates = readDateFromExcel(excelPath);
            List<String> others = readOthersFromExcel(excelPath);

            Map<String, XWPFDocument> wordDocs = new HashMap<>();
            wordDocs.put("2023", new XWPFDocument());
            wordDocs.put("2024", new XWPFDocument());
            wordDocs.put("2025", new XWPFDocument());

            for (int i = 0; i < urls.size(); i++) {
                String year = dates.get(i).split("-")[2];
                if (wordDocs.containsKey(year)) {
                    System.out.println("Processing: " + urls.get(i));
                    addToWordDoc(i + 1, wordDocs.get(year), urls.get(i), headlines.get(i), dates.get(i), others.get(i), scrapeArticle(urls.get(i)));
                    wordDocs.get(year).createParagraph().setPageBreak(true);
                }
            }

            saveWordFile(wordDocs.get("2023"), wordOutputPath2023);
            saveWordFile(wordDocs.get("2024"), wordOutputPath2024);
            saveWordFile(wordDocs.get("2025"), wordOutputPath2025);

            System.out.println("Articles saved year-wise.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Read URLs from Excel
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
        }
        return urls;
    }

    // Read Other Info
    private static List<String> readOthersFromExcel(String filePath) throws IOException {
        List<String> others = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                StringBuilder other = new StringBuilder();
                for (int i = 1; i <= 5; i++) {
                    Cell cell = row.getCell(i);
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        other.append(cell.getStringCellValue()).append(" - ");
                    }
                }
                others.add(other.toString().replaceAll(" - $", ""));
            }
        }
        return others;
    }

    // Read Headlines
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
        }
        return headlines;
    }

    // Read Dates
    private static List<String> readDateFromExcel(String filePath) throws IOException {
        List<String> dates = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                Cell cell = row.getCell(0);
                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    Date date = cell.getDateCellValue();
                    String formattedDate = new SimpleDateFormat("dd-MM-yyyy").format(date);
                    dates.add(formattedDate);
                }
            }
        }
        return dates;
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



    // Scrape Article Content
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
                    if(text.contains("UAS VISION")||
                            text.contains("an independent online news") ||
                            text.contains("Photo:") ||
                            text.contains("Photos:")){
                        continue;
                    }
                    if(text.contains("Source:")||
                            text.contains("Sources:")){
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

    // Add to Word Document
    private static void addToWordDoc(int index, XWPFDocument doc, String url, String headline, String date, String other, String content) {
        XWPFParagraph para = doc.createParagraph();
        XWPFRun urlRun = para.createRun();
        urlRun.setBold(true);
        urlRun.setFontSize(12);
        urlRun.setText(index+" . Date: " + date);
        urlRun.addCarriageReturn();
        urlRun.setText(other);
        urlRun.addCarriageReturn();
        urlRun.addCarriageReturn();
        urlRun.setText(headline);
        urlRun.addCarriageReturn();
        urlRun.addCarriageReturn();
        urlRun.setText("URL: " + url);
        urlRun.addBreak(); // Optional line break after URL


        for (String paragraph : content.split("\n\n")) {
            XWPFParagraph p = doc.createParagraph();
            XWPFRun r = p.createRun();
            r.setText(paragraph);
            r.addCarriageReturn();
        }
    }

    // Save Word File
    private static void saveWordFile(XWPFDocument doc, String outputPath) throws IOException {
        try (FileOutputStream out = new FileOutputStream(outputPath)) {
            doc.write(out);
        }
    }

}
