package com.kamar.selenium_whatsapp;

import org.apache.poi.ss.usermodel.*;
import org.apache.xmlbeans.impl.values.XmlDurationImpl;
import org.openqa.selenium.By;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

/**
 * whatsapp group members adding automation.
 * @author kamar baraka.*/

public class WhatsAppGroupAutomation {

    public static void main(String[] args)  {

        /*set the property*/
        System.setProperty("webdriver.msedge.driver", "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe");

        /*create an edge options object*/
        EdgeOptions options = new EdgeOptions();
        options.setCapability("browserVersion", "116");

        /*create an edge driver and attach the options' object*/
        WebDriver driver = new EdgeDriver(options);

        /*open the whatsapp web*/
        driver.get("https://web.whatsapp.com/");

        /*create a wait object*/
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(60));

        /*wait for the page to fully load*/
        wait.until(ExpectedConditions.jsReturnsValue("return document.readyState === 'complete'"));

        /*wait extra 30 seconds*/
        /*try{
            Thread.sleep(70000);
        }
        catch (InterruptedException exception){
            throw new RuntimeException();
        }*/

        /*wait for thirty seconds for the user to log in*/
        wait.until(ExpectedConditions.visibilityOfElementLocated
                (By.xpath("//*[@id='side']/div[1]/div/div/div[2]/div/div[1]/p")));

        /*find the search field and enter the group name*/
        WebElement searchInput = driver.findElement(By.xpath("//*[@id='side']/div[1]/div/div/div[2]/div/div[1]/p"));
        searchInput.sendKeys("Test");

        /*wait for 5 seconds for the group to appear*/
        /*wait.until(ExpectedConditions.visibilityOfElementLocated
                (By.xpath("//*[@id='pane-side']/div[1]/div/div/div[5]")));*/

        try
        {
            /*wait for browser to reload*/
//            wait.until(ExpectedConditions.jsReturnsValue("return document.readyState === 'complete"));

            try{
                Thread.sleep(5000);
            }
            catch (InterruptedException e){

                throw new RuntimeException();

            }

            /*get the group and click it*/
            WebElement group = driver.findElement(By.xpath("//*[@id='pane-side']/div[1]/div/div/div[3]"));
            group.click();
        }
        catch (StaleElementReferenceException exception){

            /*wait for a couple of seconds*//*
            try{
                Thread.sleep(2000);
            }
            catch (InterruptedException exception1){
                throw new RuntimeException();

            }

            *//*retry*//*
            WebElement group = driver.findElement(By.xpath("//*[@id='pane-side']/div[1]/div/div/div[5]/div/div/div/div[2]"));
            group.click();*/
        }

        /*wait for the text input*/
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath
                ("//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]")));

        /*find the input field for typing messages*/
        WebElement messageInput = driver.findElement
                (By.xpath("//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]"));

        /*iterate over the numbers and add*/
        Stream.of("+254748166531", "+254742585815").forEach(number -> {

            /*enter the number*/
            messageInput.sendKeys("Add me: "+ number);

            /*find the send button and click*/
            driver.findElement(By.xpath("//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[2]/button")).click();

            /*wait for some seconds before adding another*/
            try {
                Thread.sleep(30000);
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }
        });

    }

    public List<String > readExcelData(String path) throws Exception{

        /*construct a holder for our values*/
        List<String > numbers = new ArrayList<>();

        /*create an input stream of the Excel file and autoclose it*/
        try (FileInputStream inputStream = new FileInputStream(path)) {

            /*create a workbook out of the input stream*/
            Workbook workbook= WorkbookFactory.create(inputStream);

            /*get the first sheet from the workbook*/
            Sheet sheet = workbook.getSheetAt(0);

            /*iterate over the sheet to get the values and add them to the numbers */
            for (Row row :
                    sheet) {
                for (Cell cell :
                        row) {
                    /*get the vale from the cell and the value to the numbers*/
                    numbers.add(getCellValuesAsString(cell));
                }
            }
        }
        /*return the numbers*/
        return numbers;
    }

    public String getCellValuesAsString(Cell cell){

        /*check the cell type and convert it to string*/
        CellType cellType = cell.getCellType();

        return switch (cellType){
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case STRING -> cell.getStringCellValue();
            case BOOLEAN -> String.valueOf (cell.getBooleanCellValue());
            case BLANK -> "";
            default -> null;
        };
    }
}
