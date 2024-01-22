package ui;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.time.Duration;


//import com.google.common.collect.Table.Cell;


public class AutomationTesting {

	 public static void main(String[] args) throws IOException {
	        
		 System.setProperty("webdriver.chrome.driver", "C:\\browserdriver\\chromedriver.exe");
		 String excelFilePath = "C:\\Downloads\\Excel.xlsx";	
		 
		 try (FileInputStream inputStream = new FileInputStream(excelFilePath);
	             XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {

	            XSSFSheet sheet = workbook.getSheetAt(0);
	          
	            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
	                Row row = sheet.getRow(rowIndex);

	                //search term is in the second column (index 2)
	                Cell searchTermCell = row.getCell(2);

	                
	                if (searchTermCell != null) {
	                    String searchTerm = searchTermCell.getStringCellValue();

	                    // Perform Google search
	                    String googleSearchUrl = "https://www.google.com/search?q=" + searchTerm;
	                    WebDriver driver = new ChromeDriver();
	                    driver.get(googleSearchUrl);
	                    
	                    //WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
	                    //wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//*[@id=\"contents\"]/span[2]"")));

	                    // Extract suggestions
	                    List<WebElement> suggestionElements = driver.findElements(By.xpath("//*[@id=\"contents\"]/span[2]"));

	                    
	                    System.out.println("Number of suggestions: " + suggestionElements.size());

	                    
	                    String longestSuggestion = null;
	                    String lowestSuggestion = null;

	                    for (WebElement suggestionElement : suggestionElements) {
	                        String suggestion = suggestionElement.getText();
	                        if (longestSuggestion == null || suggestion.length() > longestSuggestion.length()) {
	                            longestSuggestion = suggestion;
	                        }
	                        if (lowestSuggestion == null || suggestion.length() < lowestSuggestion.length()) {
	                            lowestSuggestion = suggestion;
	                        }
	                    }

	                    row.createCell(3).setCellValue(longestSuggestion);
	                    row.createCell(4).setCellValue(lowestSuggestion);
	                   
	                    driver.quit();
	                }
	            }

	            // Save changes to the Excel file
	            try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
	                workbook.write(outputStream);
	            }

	        } catch (IOException e) {
	            e.printStackTrace();
	        }		 	  }

}
