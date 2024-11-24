import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import org.openqa.selenium.By
import org.openqa.selenium.WebDriver
import org.openqa.selenium.WebElement
import com.kms.katalon.core.webui.driver.DriverFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.ss.usermodel.Cell
import java.io.FileOutputStream
import com.kms.katalon.core.webui.driver.DriverFactory
import org.openqa.selenium.JavascriptExecutor


// Open browser and navigate to target URL
WebUI.openBrowser('')
WebUI.maximizeWindow()
WebUI.navigateToUrl('https://www.swiggy.com/city/kolkata/ganguram-maniktala-sovabazar-rest123459')

// Initialize WebDriver
WebDriver driver = DriverFactory.getWebDriver()

// Retrieve hotel name
WebElement hotelNameElement = driver.findElement(By.xpath("//div[@class='sc-jBeBSR gxFVgQ']//h1[@class='sc-aXZVg cNRZhA']"))
String hotelName = hotelNameElement.getText()
println("Hotel Name: " + hotelName)

// Retrieve additional hotel details
WebElement detailsElement = driver.findElement(By.xpath("//div[@class='sc-hBtRBD gDpXEo']"))
String rawText = detailsElement.getText()

// Extract rating, number of reviews, and outlet information
def ratingPattern = /(\d+\.\d+)/
def reviewsPattern = /\((.*?) ratings\)/
def afterOutletPattern = /Outlet\s(.+)/

def ratingMatch = (rawText =~ ratingPattern)
String rating = ratingMatch ? ratingMatch[0][0] : ""

def reviewsMatch = (rawText =~ reviewsPattern)
String numberOfReviews = reviewsMatch ? reviewsMatch[0][1] : ""

def afterOutletMatch = (rawText =~ afterOutletPattern)
String afterOutlet = afterOutletMatch ? afterOutletMatch[0][1] : ""

String combinedRating = rating + " (" + numberOfReviews + " ratings)"
println("Food outlet Rating and Reviews: " + combinedRating)
println("Area of Outlet: " + afterOutlet)

// Click dropdown elements if present
String dropdownXpath = '//div[@class="sc-fiCwlc dUWGHf"]'
boolean elementExists = true

while (elementExists) {
	try {
		if (WebUI.verifyElementPresent(findTestObject('Object Repository/dropdown_button', [('xpath') : dropdownXpath]), 2, FailureHandling.OPTIONAL)) {
			WebUI.click(findTestObject('Object Repository/dropdown_button', [('xpath') : dropdownXpath]))
			WebUI.delay(2)
		} else {
			println("Dropdown element not found.")
			elementExists = false
		}
	} catch (org.openqa.selenium.NoSuchElementException e) {
		println("Dropdown element not found.")
		elementExists = false
	}
}

// Scroll to the top of the page
JavascriptExecutor js = (JavascriptExecutor) driver
js.executeScript("window.scrollTo(0, 0);")
WebUI.delay(2)  // Optional delay to ensure the page is at the top


// Prepare to write data to Excel file
// Excel setup
String excelFilePath = 'C://Users//DELL//OneDrive//Desktop//HotelData.xlsx' // Update this path
XSSFWorkbook workbook = new XSSFWorkbook()
XSSFSheet sheet = workbook.createSheet("Hotel Data")
int rowIndex = 0

// Create header row
XSSFRow headerRow = sheet.createRow(rowIndex++)
headerRow.createCell(0).setCellValue("Hotel Name")
headerRow.createCell(1).setCellValue("Food Outlet Rating and Reviews")
headerRow.createCell(2).setCellValue("Area of Outlet")
headerRow.createCell(3).setCellValue("Dish Name")
headerRow.createCell(4).setCellValue("Dish Price")
headerRow.createCell(5).setCellValue("Dish Rating")
headerRow.createCell(6).setCellValue("Number of Reviews")

// Retrieve and print data for each dish item
List<WebElement> dishItems = driver.findElements(By.xpath("//div[@data-testid='normal-dish-item']"))
println("Number of dish items found: " + dishItems.size())

dishItems.each { WebElement dishItem ->
	String dishName = "N/A"
	String dishPrice = "N/A"
	String dishRating = "N/A"
	String dishReviews = "N/A"
	
	// Extract dish name
	try {
		WebElement dishNameElement = dishItem.findElement(By.xpath(".//div[@class='sc-aXZVg cjJTeQ sc-hIUJlX gCYyvX']"))
		dishName = dishNameElement.getText()
	} catch (Exception e) {
		println("Dish name not found for this item.")
	}
	
	// Extract dish price
	try {
		WebElement dishPriceElement = dishItem.findElement(By.xpath(".//div[@class='sc-aXZVg kCbDOU']"))
		dishPrice = dishPriceElement.getText()
	} catch (Exception e) {
		println("Dish price not found for this item.")
	}
	
	// Extract dish rating (with two possible XPaths)
	try {
		WebElement dishRatingElement = dishItem.findElement(By.xpath(".//div[@class='sc-aXZVg cFwhHc sc-krNlru borGNh']"))
		dishRating = dishRatingElement.getText()
	} catch (Exception e1) {
		try {
			WebElement dishRatingElementAlt = dishItem.findElement(By.xpath(".//div[@class='sc-aXZVg cFwhHc sc-krNlru fcKznl']"))
			dishRating = dishRatingElementAlt.getText()
		} catch (Exception e2) {
			println("Dish rating not found for this item.")
		}
	}
	
	// Extract number of reviews
	try {
		WebElement dishReviewsElement = dishItem.findElement(By.xpath(".//div[@class='sc-aXZVg jmwKWP']"))
		dishReviews = dishReviewsElement.getText()
	} catch (Exception e) {
		println("Number of reviews not found for this item.")
	}

	// Write row data to Excel
	XSSFRow row = sheet.createRow(rowIndex++)
	row.createCell(0).setCellValue(hotelName)
	row.createCell(1).setCellValue(combinedRating)
	row.createCell(2).setCellValue(afterOutlet)
	row.createCell(3).setCellValue(dishName)
	row.createCell(4).setCellValue(dishPrice)
	row.createCell(5).setCellValue(dishRating)
	row.createCell(6).setCellValue(dishReviews)
	
	// Print data for each dish
	println("Dish Name: " + dishName)
	println("Dish Price: " + dishPrice)
	println("Dish Rating: " + dishRating)
	println("Number of Reviews: " + dishReviews)
	println("----------------------")
}

println("Data extraction from dish items completed.")

// Save Excel file
FileOutputStream fileOut = new FileOutputStream(excelFilePath)
workbook.write(fileOut)
fileOut.close()
workbook.close()
println("Data saved to Excel file successfully.")

// Close the browser
//WebUI.closeBrowser()
