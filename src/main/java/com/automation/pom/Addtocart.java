package com.automation.pom;


import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class Addtocart  {

	
	
	public static WebDriver driver;
	
	
	
	@FindBy(xpath="//a[@title='Women']")
	private WebElement women;
	@FindBy(xpath="//a[@title='Casual Dresses']")
	private WebElement casualdress;
	@FindBy(xpath="//img[@title='Printed Dress']")
	private WebElement view;
	@FindBy(xpath="//*[@id='add_to_cart']/button/span")
	private WebElement cart;
	@FindBy(xpath="//a[@title='Proceed to checkout']")
	private WebElement proceed;
	@FindBy(xpath="(//a[@title='Proceed to checkout'])[2]")
	private WebElement checkout;
	@FindBy(xpath="//button[@name='processAddress']")
	private WebElement proceeds;
	@FindBy(xpath="//div[@id='uniform-cgv']")
	private WebElement checkboxs1;
	@FindBy(xpath="//button[@name='processCarrier']")
	private WebElement shipping;
	@FindBy(xpath="//*[@id='HOOK_PAYMENT']/div[1]/div/p/a")
	private WebElement cheque;
	@FindBy(xpath="//*[@id='cart_navigation']/button")
	private WebElement order;
	{
	}
	
	
		


	public Addtocart(WebDriver ldriver) {
		this.driver=ldriver;
		PageFactory.initElements(driver, this);
		}


	public static WebDriver getDriver() {
		return driver;
	}


	


	public WebElement getWomen() {
		return women;
	}


	public WebElement getCasualdress() {
		return casualdress;
	}


	public WebElement getView() {
		return view;
	}


	


	public WebElement getCart() {
		return cart;
	}


	public WebElement getProceed() {
		return proceed;
	}


	public WebElement getCheckout() {
		return checkout;
	}


	public WebElement getProceeds() {
		return proceeds;
	}


	public WebElement getCheckboxs1() {
		return checkboxs1;
	}


	public WebElement getShipping() {
		return shipping;
	}


	public WebElement getCheque() {
		return cheque;
	}


	public WebElement getOrder() {
		return order;
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	



	}


