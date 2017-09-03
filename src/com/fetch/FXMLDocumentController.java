package com.fetch;

import com.fetch.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ResourceBundle;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.PasswordField;
import javafx.scene.control.TextField;

public class FXMLDocumentController implements Initializable {
	static WebDriver driver;
	@FXML
	Label logginLinkedin;
	
	@FXML
	Button login, btnFetch;
	
	@FXML
	TextField emailLogin, fetchURL, firstName, lastName, companyNameF, locationF, domainName, comName, personalEmail, sirMam;	
	
        @FXML
        PasswordField emailPassword;
                
	@Override
	public void initialize(URL url, ResourceBundle rb) {
		//logginLinkedin.setText("sdfdsfdsf");
	}
	
	@FXML
	private void loginButtonAction(ActionEvent event) throws InterruptedException {
	    login.setDisable(true);
	    emailLogin.setDisable(true);
	    emailPassword.setDisable(true);
	    logginLinkedin.setText("Wait please...");
	    visitLinkedin();
	    Thread.sleep(5000);
	    loginLinkedinAccount(emailLogin.getText(), emailPassword.getText());
	    Thread.sleep(5000);
	    logginLinkedin.setText("You have loggedin successfully!");
	}
	
	@FXML
	private void fetchUrlAction(ActionEvent event) throws InterruptedException{
		driver.get(fetchURL.getText());
		btnFetch.setDisable(true);
		Thread.sleep(5000);
		
		String profileName = driver.findElement(By.xpath("//h1[contains(@class, 'pv-top-card-section__name')]")).getText();
		String companyName = driver.findElement(By.xpath("//h3[contains(@class, 'pv-top-card-section__company')]")).getText();
		String location = driver.findElement(By.xpath("//h3[contains(@class, 'pv-top-card-section__location')]")).getText();
		
		String[] splitedName = profileName.split("\\s+");
		
		firstName.setText(splitedName[0]);
		lastName.setText(splitedName[1]);
		companyNameF.setText(companyName);
		locationF.setText(location);
		domainName.setText(companyName.toLowerCase());
		comName.setText("com");
		btnFetch.setDisable(false);
	}
        
        @FXML
        private void sendRequestAction(ActionEvent event){
                try{
			if(driver.findElement(By.xpath("//div[contains(@class, 'pv-top-card-section__body')]/div[contains(@class, 'pv-top-card-section__actions')]//span[contains(text(), 'Connect')]")).getText().contains("Connect")){
				driver.findElement(By.xpath("//div[contains(@class, 'pv-top-card-section__body')]/div[contains(@class, 'pv-top-card-section__actions')]//span[contains(text(), 'Connect')]")).click();
				Thread.sleep(3000);
				driver.findElement(By.xpath("//div[contains(@class, 'send-invite__actions')]/button[contains(@class, 'button-primary-large')]")).click();
				Thread.sleep(1000);
				System.out.println("Request sent");
			}
		}catch(Exception e){
			System.out.println("Connect button is not listed in top section");
		}
		try{
			driver.findElement(By.xpath("//div[contains(@class, 'pv-top-card-section__right-corner-action')]")).click();
			if(driver.findElement(By.xpath("//div[contains(@class, 'pv-top-card-section__right-corner-action')]/div/div/div/ul/li[contains(@class, 'connect')]//span[contains(text(), 'Connect')]")).getText().contains("Connect")){
				driver.findElement(By.xpath("//div[contains(@class, 'pv-top-card-section__right-corner-action')]/div/div/div/ul/li[contains(@class, 'connect')]//span[contains(text(), 'Connect')]")).click();
				Thread.sleep(3000);
				driver.findElement(By.xpath("//div[contains(@class, 'send-invite__actions')]/button[contains(@class, 'button-primary-large')]")).click();
				Thread.sleep(1000);
				System.out.println("Request sent");
			}
		}catch(Exception e){
			System.out.println("Connect button not present right corner drop down menu");
		}
        }
        
        @FXML
        private void sendMessageAction(ActionEvent event){
            try{
                    
                    driver.findElement(By.xpath("//button[contains(@class, 'message-anywhere-button')]")).click();
                    Thread.sleep(1000);
                    
                    driver.findElement(By.xpath("//textarea[contains(@name, 'message')]")).sendKeys(
                                    "Hi "+firstName.getText()+","
                                    + "\n\nThank you for accepting my request. :)"
                                    + "\n\nJust wanted to ask you something related to "+companyNameF.getText()+"? May I please?"
                                    + "\n\nI am sorry if you find anything wrong here. please accept my apologies."
                                    + "\n\nSincerely,\nYour name");
                    Thread.sleep(3000);
		
                    driver.findElement(By.xpath("//button[contains(@class, 'msg-messaging-form__send-button')]")).click();
                   
		}catch(Exception e){
			System.out.println("Message button is not present");
		}
        }
	
	@FXML
	private void exportToExcelF(ActionEvent event) throws EncryptedDocumentException, InvalidFormatException, IOException{
		exportToExcel(firstName.getText(), lastName.getText(), companyNameF.getText(), locationF.getText(), domainName.getText(), comName.getText(), personalEmail.getText(), sirMam.getText());
		
		firstName.setText("");
		lastName.setText("");
		companyNameF.setText("");
		locationF.setText("");
		domainName.setText("");
		comName.setText("");
		personalEmail.setText("");
		sirMam.setText("");
	}
	
	
	public static void loginLinkedinAccount(String email, String pass){
		driver.findElement(By.xpath("//input[contains(@class, 'login-email')]")).sendKeys(email);
		driver.findElement(By.xpath("//input[contains(@class, 'login-password')]")).sendKeys(pass);
		driver.findElement(By.xpath("//input[contains(@id, 'login-submit')]")).click();
	}

	public static void visitLinkedin(){
		ChromeOptions chromeOptions = new ChromeOptions();
		File chromeDriver = new File("C:\\chromedriver.exe");
		System.setProperty("webdriver.chrome.driver", chromeDriver.getAbsolutePath());
		chromeOptions.addArguments("--headless");
		chromeOptions.addArguments("--disable-gpu");
		chromeOptions.addArguments("--window-size=1024, 768");
		
		driver = new ChromeDriver(chromeOptions);
		driver.get("https://www.linkedin.com/");
	}
	
	
	public static void exportToExcel(String firstName, String lastName, String companyName, String location, String domainName, String comName, String personalEmail, String sirMam) throws EncryptedDocumentException, InvalidFormatException, IOException{
		XSSFWorkbook wb;
		Sheet sheet;
		Row row = null;
		
		if(!new File("F://Book_list.csv").isFile()) { 
//		    System.out.println("file not present");
			wb = new XSSFWorkbook();
	        sheet = (XSSFSheet) wb.createSheet();
	        row = sheet.createRow(0);
	        sheet = wb.createSheet();
	        
	        row.createCell(0).setCellValue("firstName");
	        row.createCell(1).setCellValue("lastName");
	        row.createCell(2).setCellValue("companyName");
	        row.createCell(3).setCellValue("companyLocation");
	        row.createCell(4).setCellValue("domainName");
	        row.createCell(5).setCellValue("comName");
	        row.createCell(6).setCellValue("emailLogin");
	        row.createCell(7).setCellValue("passwordLogin");
	        row.createCell(8).setCellValue("personalEmail");
	        row.createCell(9).setCellValue("Email Status");
	        row.createCell(10).setCellValue("Sir/Mam");
		}else{
			InputStream inp = new FileInputStream("F://Book_list.csv");
			wb = (XSSFWorkbook) WorkbookFactory.create(inp);
		    sheet = wb.getSheetAt(0);
		    int rowNum = sheet.getPhysicalNumberOfRows();
		    //System.out.println(rowNum);
		    row = sheet.createRow(rowNum++);
		    
		    row.createCell(0).setCellValue(firstName);
	        row.createCell(1).setCellValue(lastName);
	        row.createCell(2).setCellValue(companyName);
	        row.createCell(3).setCellValue(location);
	        row.createCell(4).setCellValue(domainName);
	        row.createCell(5).setCellValue(comName);
	        row.createCell(6).setCellValue("");
	        row.createCell(7).setCellValue("");
	        if(personalEmail.length() != 0){
	        	row.createCell(8).setCellValue(personalEmail);
	        }else{
	        	row.createCell(8).setCellValue("N");
	        }
	        row.createCell(9).setCellValue("No status");
	        row.createCell(10).setCellValue(sirMam);
		}
      
      
	    FileOutputStream fileOut = new FileOutputStream("F://Book_list.csv");
	    wb.write(fileOut);
	    fileOut.close();
	}
	
}
