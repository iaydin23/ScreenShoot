package pages;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import utilities.Driver;

public class AmazonPage {

    public AmazonPage(){
        PageFactory.initElements(Driver.getDriver(),this);

    }


    @FindBy(id="twotabsearchtextbox")
    public WebElement aramakutusu;

    @FindBy (xpath = "//div[@class='a-section a-spacing-small a-spacing-top-small']")
    public WebElement aramaSonucWE;

    @FindBy (xpath="(//div[@class='a-section aok-relative s-image-square-aspect'])[1]")
    public WebElement ilkUrun;

    @FindBy (xpath="(//div[@class='a-section a-spacing-base'])[1]")
    public WebElement ilkUrunDetay;


}
