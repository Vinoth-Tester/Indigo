package org.utilityclass;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Set;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriver.Navigation;
import org.openqa.selenium.WebDriver.Options;
import org.openqa.selenium.WebDriver.TargetLocator;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class UtilityClass {

	public static WebDriver driver;

	// Browser launch

	public WebDriver browserLaunch() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		return driver;
	}

	// Methods in WebDriver Interface

	public void loadUrl(String url) {
		driver.get(url);
	}

	public Navigation navigateDriver() {
		Navigation navigate = driver.navigate();
		return navigate;
	}

	public String getPageUrl() {
		String currentUrl = driver.getCurrentUrl();
		return currentUrl;
	}

	public void quitBrowser() {
		driver.quit();
	}

	public void closeBrowser() {
		driver.close();
	}

	public Options manageClass() {
		Options manage = driver.manage();
		return manage;
	}

	public TargetLocator switchToElement() {
		TargetLocator switchTo = driver.switchTo();
		return switchTo;
	}

	// navigation commands

	public void goToUrl(String url) {
		driver.navigate().to(url);
	}

	public void forwardToUrl(String url) {
		driver.navigate().forward();
	}

	public void refreshUrl(String url) {
		driver.navigate().refresh();
	}

	public void backToUrl(String url) {
		driver.navigate().back();
	}

	// Methods in WebElement Interface

	public void insertText(WebElement tb, String txt) {
		tb.sendKeys(txt);
	}

	public void clearText(WebElement tb) {
		tb.clear();
	}

	public void btnClick(WebElement tb) {
		tb.click();
	}

	public String getElementText(WebElement tb) {
		String text = tb.getText();
		return text;
	}

	public String getWindowId() {
		String windowHandle = driver.getWindowHandle();
		return windowHandle;
	}

	public Set<String> getWindowsIds() {
		Set<String> windowHandles = driver.getWindowHandles();
		return windowHandles;
	}

	public void switchToParentWindow() {
		String parentWindow = driver.getWindowHandle();
		driver.switchTo().window(parentWindow);
	}

	public void switchToChildWindow() {
		String parentWindow = driver.getWindowHandle();
		Set<String> windowHandles = driver.getWindowHandles();
		for (String childWindow : windowHandles) {
			if (!parentWindow.equals(childWindow)) {
				driver.switchTo().window(childWindow);
			}
		}
	}

	public WebElement FindElementById(String locator) {
		WebElement element = driver.findElement(By.id(locator));
		return element;
	}

	public WebElement FindElementByName(String locator) {
		WebElement element = driver.findElement(By.name(locator));
		return element;
	}

	public WebElement FindElementByClassName(String locator) {
		WebElement element = driver.findElement(By.className(locator));
		return element;
	}

	public WebElement FindElementByXPath(String locator) {
		WebElement element = driver.findElement(By.xpath(locator));
		return element;
	}

	public List<WebElement> FindElementsByClassName(String locators) {
		List<WebElement> elements = driver.findElements(By.className(locators));
		return elements;
	}

	public List<WebElement> FindElementsById(String locators) {
		List<WebElement> elements = driver.findElements(By.id(locators));
		return elements;
	}

	public List<WebElement> FindElementsByName(String locators) {
		List<WebElement> elements = driver.findElements(By.name(locators));
		return elements;
	}

	public List<WebElement> FindElementsByXPath(String locators) {
		List<WebElement> elements = driver.findElements(By.xpath(locators));
		return elements;
	}

	public String getPageTitle() {
		String title = driver.getTitle();
		return title;
	}

	public String getAttributeValue(WebElement gav, String name) {
		String text = gav.getAttribute(name);
		return text;
	}

	public boolean selected(WebElement se) {
		boolean selected = se.isSelected();
		return selected;
	}

	public boolean displayed(WebElement se) {
		boolean displayed = se.isSelected();
		return displayed;
	}

	public boolean enabled(WebElement se) {
		boolean enabled = se.isSelected();
		return enabled;
	}

	// Classes in Actions Class

	public void mouseOverAction(WebElement target) {
		Actions a = new Actions(driver);
		a.moveToElement(target).perform();
	}

	public void dragAndDropAction(WebElement src, WebElement dest) {
		Actions a = new Actions(driver);
		a.dragAndDrop(src, dest).perform();
	}

	public void rightClick(WebElement target) {
		Actions a = new Actions(driver);
		a.contextClick().perform();
	}

	public void clickTwice(WebElement target) {
		Actions a = new Actions(driver);
		a.doubleClick().perform();
	}

	public void performAction() {
		Actions a = new Actions(driver);
		a.perform();
	}

	// Methods in Select Class

	public void selectOptionsByIndex(WebElement sec, int index) {
		Select s = new Select(sec);
		s.selectByIndex(index);
	}

	public void selectOptionsByValue(WebElement sec, String value) {
		Select s = new Select(sec);
		s.selectByValue(value);
	}

	public void selectOptionsByText(WebElement sec, String text) {
		Select s = new Select(sec);
		s.selectByVisibleText(text);
	}

	public void deSelectOptionsByIndex(WebElement sec, int index) {
		Select s = new Select(sec);
		s.deselectByIndex(index);
	}

	public void deSelectOptionsByValue(WebElement sec, String value) {
		Select s = new Select(sec);
		s.deselectByValue(value);
	}

	public void deSelectOptionsByText(WebElement sec, String text) {
		Select s = new Select(sec);
		s.deselectByVisibleText(text);
	}

	public void deSelectAllOptions(WebElement sec) {
		Select s = new Select(sec);
		s.deselectAll();
	}

	public List<WebElement> getOptionFromDropDown(WebElement sec) {
		Select s = new Select(sec);
		List<WebElement> options = s.getOptions();
		return options;
	}

	public WebElement getFirstSelectedOptionFromDropDown(WebElement sec) {
		Select s = new Select(sec);
		WebElement firstSelectedOption = s.getFirstSelectedOption();
		return firstSelectedOption;
	}

	public List<WebElement> getAll(WebElement sec) {
		Select s = new Select(sec);
		List<WebElement> allSelectedOptions = s.getAllSelectedOptions();
		return allSelectedOptions;
	}

	// Screenshot

	private void screenShotElement(String path, WebElement element) throws IOException {

		File screenshotAs = element.getScreenshotAs(OutputType.FILE);
		FileUtils.copyFile(screenshotAs, new File(path));

	}

	private void screenShot(String path) throws IOException {
		TakesScreenshot shot = (TakesScreenshot) driver;
		File screenshotAs = shot.getScreenshotAs(OutputType.FILE);
		FileUtils.copyFile(screenshotAs, new File(path));
	}

	// Methods in Alert Interface

	public void acceptAlert() {
		Alert al = driver.switchTo().alert();
		al.accept();
	}

	public void dismissAlert() {
		Alert al = driver.switchTo().alert();
		al.dismiss();
	}

	public void SendTextToAlert(String text) {
		Alert al = driver.switchTo().alert();
		al.sendKeys(text);
	}

	public String getAlertText() {
		Alert al = driver.switchTo().alert();
		String text = al.getText();
		return text;
	}

	// Methods in Frame

	public void switchToFramesById(String id) {
		switchToElement().frame(id);
	}

	public void switchToFramesByName(String name) {
		switchToElement().frame(name);
	}

	public void switchToFramesByIndex(int index) {
		switchToElement().frame(index);
	}

	public void switchToFramesByWebElement(WebElement ref) {
		switchToElement().frame(ref);
	}

	// ExcelIntegration

	public String GetDataFromExcel(String path, String sheet, int rowindex, int cellindex) throws IOException {

		File excelLoc = new File(path);
		String value = null;
		FileInputStream stream = new FileInputStream(excelLoc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet(sheet);
		Row r = s.getRow(rowindex);
		Cell c = r.getCell(cellindex);
		int type = c.getCellType();
		if (type == 1) {
			value = c.getStringCellValue();
		}
		if (type == 0) {
			if (DateUtil.isCellDateFormatted(c)) {
				Date d = c.getDateCellValue();
				SimpleDateFormat sim = new SimpleDateFormat("dd-MM-yyyy");
				value = sim.format(d);
			} else {
				double cellValue = c.getNumericCellValue();
				long l = (long) cellValue;
				value = String.valueOf(l);

			}
		}
		return value;

	}

	public void writeDataToExcel(String data, String path, String sheet, int rowindex, int cellindex)
			throws IOException {

		File excelLoc = new File(path);
		Workbook w = new XSSFWorkbook();
		Sheet s = w.createSheet(sheet);
		Row r = s.createRow(rowindex);
		Cell c = r.createCell(cellindex);

		c.setCellValue(data);
		FileOutputStream ostream = new FileOutputStream(excelLoc);
		w.write(ostream);

	}

	public void updateDataToExcel(String olddata, String newdata, String path, String sheet, int rowindex,
			int cellindex) throws IOException {

		File excelLoc = new File(path);
		FileInputStream stream = new FileInputStream(excelLoc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet(sheet);
		Row r = s.getRow(rowindex);
		Cell c = r.getCell(cellindex);
		String value = c.getStringCellValue();

		if (value.equals(olddata)) {
			c.setCellValue(newdata);
		}
		FileOutputStream ostream = new FileOutputStream(excelLoc);
		w.write(ostream);

	}

	public void WriteDataInCell(String data, String path, String sheet, int rowindex, int cellindex)
			throws IOException {
		File excelLoc = new File(path);
		FileInputStream stream = new FileInputStream(excelLoc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet(sheet);
		Row r = s.getRow(rowindex);
		Cell c = r.createCell(cellindex);
		c.setCellValue(data);
		FileOutputStream ostream = new FileOutputStream(excelLoc);
		w.write(ostream);

	}
}
