package org.web.parser;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class WebParser {

    static class Product {
        String productCategory;
        String name;
        String rate;

        public Product(String productCategory, String name, String rate) {
            this.productCategory = productCategory;
            this.name = name;
            this.rate = rate;
        }
    }

    public static void main(String[] args) throws Exception {
        // Setup WebDriverManager for Chrome
        WebDriverManager.firefoxdriver().setup();
        WebDriver driver = new FirefoxDriver();

        driver.get("https://sivakasiskmcrackers.in/"); // Replace with your actual URL
        driver.manage().window().maximize();

        scrollToEnd((JavascriptExecutor) driver);

        List<Product> productList = new ArrayList<>();

        List<WebElement> products = driver.findElements(By.xpath("//div[@class='container-fluid product']/div[@class='row']/div"));
        String productCategory = "N/A";

        for (WebElement product : products) {
            String productName = "N/A";
            String rate = "N/A";

            try {
                WebElement productCategoryElement = product.findElement(By.xpath(".//h1[@class='catbgclr heading6 text-center roboto mb-0']"));
                if (productCategoryElement != null) {
                    productCategory = productCategoryElement.getText().trim();
                }
            } catch (Exception ignored) {
            }
            System.out.println("Product Category: " + productCategory);
            List<WebElement> productsInfo = product.findElements(By.xpath(".//div[@class='row pb-2 product-form']/div"));
            System.out.println("Product Info Size: " + productsInfo.size());

            for (WebElement productInfo : productsInfo) {
                try {
                    productName = productInfo.findElement(By.xpath(".//div[contains(@class, 'producttext')]")).getText().trim();
                } catch (Exception ignored) {
                }
                System.out.println("Product Name: " + productName);
                try {
                    rate = productInfo.findElement(By.xpath(".//span[@class='rate']")).getText().trim();
                } catch (Exception ignored) {
                }
                System.out.println("Rate: " + rate);
                System.out.println("--------------------------------------------------");

                productList.add(new Product(productCategory, productName, rate));
            }

            System.out.println("==================================================");
        }
        driver.quit();
        exportReport(productList);
    }

    private static void scrollToEnd(JavascriptExecutor driver) throws InterruptedException {
        JavascriptExecutor js = driver;

        // --- Scroll until all products loaded ---
        int lastHeight = ((Number) js.executeScript("return document.body.scrollHeight")).intValue();

        while (true) {
            js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
            Thread.sleep(2000); // wait for AJAX to load products

            int newHeight = ((Number) js.executeScript("return document.body.scrollHeight")).intValue();
            if (newHeight == lastHeight) {
                break; // reached end
            }
            lastHeight = newHeight;
        }
    }

    private static void exportReport(List<Product> productList) throws IOException {
        // --- Excel Report ---
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Products");

        int rowCount = 0;
        Row header = sheet.createRow(rowCount++);
        header.createCell(0).setCellValue("Product Category");
        header.createCell(1).setCellValue("Product Name");
        header.createCell(2).setCellValue("Rate");

        for (Product p : productList) {
            Row row = sheet.createRow(rowCount++);
            row.createCell(0).setCellValue(p.productCategory);
            row.createCell(1).setCellValue(p.name);
            row.createCell(2).setCellValue(p.rate);
        }

        try (FileOutputStream out = new FileOutputStream("Products.xlsx")) {
            workbook.write(out);
        }
        workbook.close();

        // --- Word Report ---
        XWPFDocument doc = new XWPFDocument();
        XWPFTable table = doc.createTable(productList.size() + 1, 3);

        // Header row
        table.getRow(0).getCell(0).setText("Category");
        table.getRow(0).getCell(1).setText("Product Name");
        table.getRow(0).getCell(2).setText("Rate");

        int rowIdx = 1;
        for (Product p : productList) {
            table.getRow(rowIdx).getCell(0).setText(p.productCategory);
            table.getRow(rowIdx).getCell(1).setText(p.name);
            table.getRow(rowIdx).getCell(2).setText(p.rate);
            rowIdx++;
        }

        try (FileOutputStream out = new FileOutputStream("Products.docx")) {
            doc.write(out);
        }
        doc.close();

        System.out.println("Reports generated: Products.xlsx and Products.docx");
    }
}