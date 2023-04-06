package org.isp1;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Day1 {
	
	public static void main(String[] args) {
		
		WebDriverManager.chromedriver().setup();
		ChromeOptions ops = new ChromeOptions();
		ops.addArguments("--remote-allow-origins=*");
		WebDriver driver=new ChromeDriver(ops);
		 driver.get("https://onboard.zupain.com/");
		
		
	}
	
	
	

}
	