package files;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.time.Duration;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class Project{

	public static void main(String[] args) {
		//drive part
		WebDriverManager.firefoxdriver().setup();
		WebDriver driver = new FirefoxDriver();

	try {
//		#Step 01 - Get specific value to specific cell to excel file
//		    input excel file
			FileInputStream excelFile = new FileInputStream("D:\\Q\\finalexcel.xlsx");
//			Create new workbook
			Workbook workbook = new XSSFWorkbook(excelFile);
//			Get the sheet, insert index number
			Sheet sheet = workbook.getSheetAt(0);
//		   Insert row index number
			Row row = sheet.getRow(11);
//		   Insert clu index number
			Cell cell = row.getCell(2);
			// Read  cell value
			String cellValue = cell.getStringCellValue();
//       #Step 02 -Open Google and perform a search
			driver.get("https://www.google.com");
//			locator find
			WebElement searchBox = driver.findElement(By.name("q"));
//			value pass search filed
			searchBox.sendKeys(cellValue);

//       #Step 03 - Suggestion value find and conditions
		// Wait for the suggestions dropdown to appear
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(5));
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("sbct")));

		// Get all suggestion elements within the dropdown
		List<WebElement> suggestionElements = driver.findElements(By.className("sbct"));

		String longestSuggestion = findLongestNonEmptySuggestion(suggestionElements);
		String shortestSuggestion = findShortestNonEmptySuggestion(suggestionElements);
//
		// Print the longest and shortest suggestions
		System.out.println("Longest suggestion: " + longestSuggestion);
		System.out.println("Shortest suggestion: " + shortestSuggestion);

// Step 04 -Excel write part
		// insert actual sheet name
		Sheet outputSheet = workbook.getSheet("Saturday");

		// Decide row number where you write .
		Row outputRow = outputSheet.getRow(11);

		// Decide column number where you write .
		Cell longestCell = outputRow.createCell(3);
		longestCell.setCellValue(longestSuggestion);


//       Decide column number where you write .
		Cell shortestCell = outputRow.createCell(4);
		shortestCell.setCellValue(shortestSuggestion);


		// Step 5: Save the changes to the Excel file
		FileOutputStream outputStream = new FileOutputStream("D:\\Q\\finalexcel.xlsx");
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
		System.out.println("Value written successfully.");

	} catch (Exception e) {
		e.printStackTrace();
	} finally {
		// Quit the WebDriver
		driver.quit();
	}
	}

	private static String findLongestNonEmptySuggestion(List<WebElement> suggestionElements) {
		String longestSuggestion = null;

		for (WebElement suggestion : suggestionElements) {
			String suggestionText = suggestion.getText().trim(); // Trim to remove leading/trailing spaces

			// Skip empty or zero-length suggestions
			if (suggestionText.isEmpty()) {
				continue;
			}

			if (longestSuggestion == null || isLongerSuggestion(suggestionText, longestSuggestion)) {
				longestSuggestion = suggestionText;
			}
		}

		return longestSuggestion;
	}

	private static String findShortestNonEmptySuggestion(List<WebElement> suggestionElements) {
		String shortestSuggestion = null;

		for (WebElement suggestion : suggestionElements) {
			String suggestionText = suggestion.getText().trim(); // Trim to remove leading/trailing spaces

			// Skip empty or zero-length suggestions
			if (suggestionText.isEmpty()) {
				continue;
			}

			if (shortestSuggestion == null || isShorterSuggestion(suggestionText, shortestSuggestion)) {
				shortestSuggestion = suggestionText;
			}
		}

		return shortestSuggestion;
	}

	private static boolean isLongerSuggestion(String suggestion, String currentLongest) {
		return suggestion.length() > currentLongest.length();
	}

	private static boolean isShorterSuggestion(String suggestion, String currentShortest) {
		return suggestion.length() < currentShortest.length();
	}
}
