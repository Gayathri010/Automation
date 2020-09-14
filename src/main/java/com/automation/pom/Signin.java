package com.automation.pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class Signin {
	
	public static WebDriver driver;
	@FindBy(id="email")
	private WebElement email;
	
	@FindBy(id="passwd")
	private WebElement password;

	@FindBy(id="SubmitLogin")
	private WebElement Submit;

	

	public Signin(WebDriver ldriver) {
		this.driver=ldriver;
		PageFactory.initElements(driver, this);	}



	public static WebDriver getDriver() {
		return driver;
	}



	public WebElement getEmail() {
		return email;
	}



	public WebElement getPassword() {
		return password;
	}



	public WebElement getSubmit() {
		return Submit;
	}

	

	
	
}
