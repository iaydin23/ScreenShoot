package tests.karısıkSoruCozumleri;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.testng.annotations.Test;
import pages.AmazonPage;
import utilities.ConfigReader;
import utilities.Driver;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

public class AmazonSearchTestDatToExcelWithScreenShots {

    @Test
    public void amazonSearchTest() throws IOException {

        Driver.getDriver().get(ConfigReader.getProperty("amazonUrl"));
        Driver.getDriver().navigate().refresh();

        AmazonPage amazonPage = new AmazonPage();

        String filePath = "src/resources/amazonsearch.xlsx";
        FileInputStream fis = new FileInputStream(filePath);

        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        TakesScreenshot ts = (TakesScreenshot) Driver.getDriver();

        String aramaSonucu;
        File file;
        String arananKelime;
        int lastRowinSheet1 = sheet.getLastRowNum();
        Path imagePath;
        byte[] imageContent;
        XSSFClientAnchor anchor;
        XSSFDrawing drawingPatriarch;
        for (int i = 1; i < lastRowinSheet1; i++) {
            arananKelime = sheet.getRow(i).getCell(0).toString();
            amazonPage.aramakutusu.sendKeys(arananKelime, Keys.ENTER);
            aramaSonucu = amazonPage.aramaSonucWE.getText();
            sheet.getRow(i).createCell(1).setCellValue(aramaSonucu);
            sheet.getRow(i).createCell(2).setCellValue("target/screen-shots/SS" + arananKelime + ".jpeg");

            file = ts.getScreenshotAs(OutputType.FILE);
            FileUtils.copyFile(file, new File("target/screen-shots/SS" + arananKelime + ".jpeg"));

            imagePath = Path.of("target/screen-shots/SS" + arananKelime + ".jpeg");
            imageContent = Files.readAllBytes(imagePath);
            int pictureIndex = workbook.addPicture(imageContent, Workbook.PICTURE_TYPE_JPEG);
            anchor = new XSSFClientAnchor(0, 0, 0, 0, 3, (i), 5, (i + 1));

            drawingPatriarch = sheet.createDrawingPatriarch();
            drawingPatriarch.createPicture(anchor, pictureIndex);

            amazonPage.aramakutusu.clear();
        }
        FileOutputStream fos = new FileOutputStream(filePath);
        workbook.write(fos);

        workbook.close();
        fos.close();
        fis.close();
        Driver.quitDriver();
    }


}
