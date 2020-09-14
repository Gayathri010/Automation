package com.automation.pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class Homepage {

	
	public static WebDriver driver;
	
	@FindBy(xpath="//a[@class='login']")
	private WebElement signin;
	
	
	public Homepage(WebDriver ldriver) {
		this.driver=ldriver;
		PageFactory.initElements(driver, this);
	}


	public static WebDriver getDriver() {
		return driver;
	}


	public WebElement getSignin() {
		return signin;
	}
	
}
