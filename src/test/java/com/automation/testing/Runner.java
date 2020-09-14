package com.automation.testing;

import org.junit.Test;
import org.openqa.selenium.WebDriver;

import com.automation.baseclass.Baseclass;
import com.automation.pom.Addtocart;
import com.automation.pom.Homepage;
import com.automation.pom.Signin;

public class Runner extends Baseclass {

	public static WebDriver driver;

	//public static void main(String[] args) throws InterruptedException {
	
	@Test
	public void test() throws InterruptedException {

		driver = Browser("chrome");
		getUrl("http://automationpractice.com/index.php");
		Homepage hp = new Homepage(driver);
		clickOnElement(hp.getSignin());

		Signin sp = new Signin(driver);
		inputValueElement(sp.getEmail(), "gayudevi0217@gmail.com");
		inputValueElement(sp.getPassword(), "murugan");
		clickOnElement(sp.getSubmit());

		Addtocart ac = new Addtocart(driver);
		actionMethod(ac.getWomen());
		clickOnElement(ac.getCasualdress());
		clickOnElement(ac.getView());
		Thread.sleep(3000);
		driver.switchTo().frame(0);
		clickOnElement(ac.getCart());
		clickOnElement(ac.getProceed());
		clickOnElement(ac.getCheckout());
		clickOnElement(ac.getProceeds());
		clickOnElement(ac.getCheckboxs1());
		clickOnElement(ac.getShipping());
		clickOnElement(ac.getCheque());
		clickOnElement(ac.getOrder());

	}

}
