import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.webui.driver.DriverFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFCell
import java.io.FileOutputStream
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.WebDriver
import org.openqa.selenium.WebElement
import org.openqa.selenium.By
import org.openqa.selenium.NoSuchElementException
import org.openqa.selenium.support.ui.WebDriverWait
import org.openqa.selenium.support.ui.ExpectedConditions

// Initialize WebDriver and navigate to the target URL
WebUI.openBrowser('')
WebUI.navigateToUrl('https://www.swiggy.com/city/kolkata/best-restaurants')
WebUI.maximizeWindow()
WebDriver driver = DriverFactory.getWebDriver()

// Click "Show more" button until no more results are displayed
String xpath = "//div[contains(text(),'Show more')]"
boolean elementExists = true
int retryCount = 0

while (elementExists && retryCount < 10) {
	try {
		WebElement showMoreButton = driver.findElement(By.xpath(xpath))
		showMoreButton.click()
		WebUI.delay(2)
		retryCount++
	} catch (NoSuchElementException e) {
		elementExists = false
	}
}

// Scroll to the top of the page
JavascriptExecutor js = (JavascriptExecutor) driver
js.executeScript("window.scrollTo(0, 0);")

// Excel setup
String excelFilePath = 'C://Users//DELL//OneDrive//Desktop//cityOutletData.xlsx' // Update this path
XSSFWorkbook workbook = new XSSFWorkbook()
XSSFSheet sheet = workbook.createSheet("Outlet Data")
String[] headers = ["Outlet Name", "Rating", "Food Type", "Area", "Link"]
XSSFRow headerRow = sheet.createRow(0)
headers.eachWithIndex { header, index ->
	XSSFCell cell = headerRow.createCell(index)
	cell.setCellValue(header)
}

// Wait setup
WebDriverWait wait = new WebDriverWait(driver, 10)

// Locate grid cards and extract data
List<WebElement> gridCards = driver.findElements(By.xpath('//div[@class="styled__StyledRestaurantGridCard-sc-fcg6mi-0 lgOeYp"]'))
println("Number of outlets found: " + gridCards.size())

gridCards.eachWithIndex { WebElement gridCard, int rowIndex ->
	String outletName = "N/A"
	String rating = "N/A"
	String foodType = "N/A"
	String area = "N/A"
	String link = "N/A"

	try {
		outletName = gridCard.findElement(By.xpath('.//div[@class="sc-beySbM iKLEMo"]')).getText()
	} catch (NoSuchElementException e) {
		println("Outlet name not found for this card.")
	}

	try {
		rating = gridCard.findElement(By.xpath('.//div[@class="sc-beySbM brneVe"]/span[@class="sc-beySbM jAtOQO"]')).getText().replaceAll(" •", "")
	} catch (NoSuchElementException e) {
		println("Rating not found for this card.")
	}

	try {
		foodType = gridCard.findElement(By.xpath('.//div[@class="sw-restaurant-card-descriptions-container"]/div[@class="sc-beySbM bRTXBF"][1]')).getText()
	} catch (NoSuchElementException e) {
		println("Food type not found for this card.")
	}

	try {
		area = gridCard.findElement(By.xpath('.//div[@class="sw-restaurant-card-descriptions-container"]/div[@class="sc-beySbM bRTXBF"][2]')).getText()
	} catch (NoSuchElementException e) {
		println("Area not found for this card.")
	}

try {
    // Locate the anchor (`<a>`) tag and get the `href` attribute
    WebElement linkElement = gridCard.findElement(By.xpath('.//ancestor::a[@class="RestaurantList__RestaurantAnchor-sc-1d3nl43-3 kcEtBq"]'))
    link = linkElement.getAttribute("href")
    println("Link found: " + link)
} catch (NoSuchElementException e) {
    println("Link not found for this card.")
} catch (Exception e) {
    println("Error while retrieving link: " + e.getMessage())
}
	
	// Write extracted data to the Excel sheet
	XSSFRow row = sheet.createRow(rowIndex + 1)
	row.createCell(0).setCellValue(outletName)
	row.createCell(1).setCellValue(rating)
	row.createCell(2).setCellValue(foodType)
	row.createCell(3).setCellValue(area)
	row.createCell(4).setCellValue(link)
	
	println("Outlet Name: " + outletName)
	println("Rating: " + rating)
	println("Food Type: " + foodType)
	println("Area: " + area)
	println("Link: " + link)
	println("----------------------")
}

// Save the Excel file
FileOutputStream fileOut = new FileOutputStream(excelFilePath)
workbook.write(fileOut)
fileOut.close()
workbook.close()
println("Data extraction completed and saved to Excel file.")

// Close the browser
WebUI.closeBrowser()
