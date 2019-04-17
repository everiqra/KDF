import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;


public class KDF_Driver3 {
	// Global Variables and object 
	static WebDriver myD;
	static String[][] xTC,xTS;
	static int xTC_R, xTS_R;
	static String xlPath= "C:\\Users\\WAQAS\\Slot_nov9\\KDF5_YTHardCoded.xls";

	public static void main(String[] args) throws Exception{
		String vTCID, VTC_Exe;
		String vTSID, VTS_StepNum, VKeyword, VDATA , VEID;

		// 1. read the i/p KDF KDF1
		xTC=readXL(xlPath,"TestCases") ;
		xTS=readXL(xlPath,"TestSteps") ;

		xTC_R = xTC.length;
		xTS_R = xTS.length;

		System.out.println("Test cases row "+xTC.length );
		System.out.println("Test step row "+xTS.length );

		//identify Executable Test cases 
		for (int i =1; i <xTC_R; i++){
			System.out.println("**************TC start********************" );
			vTCID = xTC[i][0] ; 
			VTC_Exe =xTC [i][2];

			if (VTC_Exe.equals("Y")){
				System.out.println("TCID Ready for execution "+vTCID );

				// Go through the each row in test step sheet
				for (int j =1; j <xTS_R; j++){

					vTSID =  xTS[j][0] ;
					if(vTCID.equals(vTSID)){ // Executed only matching TCID's
						//System.out.println("************TS start**************" );


						VTS_StepNum = xTS[j][1] ;
						VKeyword = xTS[j][3] ;
						VDATA= xTS[j][4]  ; 
						VEID= xTS[j][5] ;
						System.out.println("##"+ VTS_StepNum);
						System.out.println("KW :::"+VKeyword );
						System.out.println("TD :::" +VDATA );
						System.out.println("EID :::"+VEID );

						//Execute the test step 
				try{		
					 executeKW(VKeyword,VDATA,VEID );
					 System.out.println("step>>>>>pass");
				}catch (Exception e){
					System.out.println("step>>>>>pass:::"+ e);
				}


						//System.out.println("********TS END*******" );

					}	else {
						// Skipping the irrelevant step
					}
				}

			}else {
				//System.out.println("**********************************" );
				System.out.println(">>>>>>>>>>>>TCID Skipped "+vTCID );

			}
			System.out.println("**************TC End********************" );

		}
	}

	
	// Reusable Components 
	public static void  executeKW(String fKW, String fTD, String fEID) {
		//Execute the test step 
		if( fKW.equals("openBrowser")){
			openBrowser( fTD );
		}
		if( fKW.equals("navigateBrowser")){
			navigateBrowser( fTD );
		}
		if( fKW.equals("typeText")){
			typeText( fTD, fEID );
		}
		if( fKW.equals("clickElement")){
			clickElement(fEID);
		}
		if( fKW.equals("verifyText")){
			verifyText( fTD, fEID  );
		}
		if( fKW.equals("closeBrowser")){
			closeBrowser(  );
		}
		
	}
	
	
	public static String[][] readXL(String fPath, String fSheet) throws Exception{
		//input: XL Path and XL Sheet name
		// Output
		String [][] xData;
		int xRows, xCols;
		DataFormatter dataFormatter = new DataFormatter();
		String cellValue;
		File myxl= new File (fPath);
		FileInputStream myStream = new FileInputStream(myxl); 
		HSSFWorkbook myWB = new HSSFWorkbook(myStream);
		HSSFSheet mySheet = myWB.getSheet(fSheet);
		xRows = mySheet.getLastRowNum()+1;
		xCols = mySheet.getRow(0).getLastCellNum();
		xData = new String[xRows][xCols];

		System.out.println("number of row "+ xRows);
		System.out.println("number of colms "+ xCols);
		System.out.println("~~~~~~~~~~~~~~~~~Test data below ~~~~~~~~~~~~~~");


		for(int i = 0; i< xRows; i++){
			HSSFRow row = mySheet.getRow(i);
			for (int j=0; j< xCols; j++){
				cellValue = "_";
				cellValue= dataFormatter.formatCellValue(row.getCell(j));
				if (cellValue!=null){
					xData[i][j]= cellValue;
				}System.out.print(cellValue);
				System.out.print("||||");
			}System.out.println("");

		}
		myxl = null; // Memory get released
		return xData;	
	}
	// Method to write into xL
	public static void writeXL(String fPath, String fSheet, String[][] xData) throws Exception{
		File outFile= new File (fPath);
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet osheet = wb.createSheet(fSheet);
		int xR_TS = xData.length;
		int xC_TS = xData[0].length;

		for(int myrow = 0; myrow< xR_TS; myrow++){
			HSSFRow row = osheet.createRow(myrow);
			for (int mycol=0; mycol< xC_TS; mycol++){
				HSSFCell cell = row.createCell(mycol);
				//cell.setCellType(HSSFCell.Cell_TYPE_STRING);
				cell.setCellValue(xData[myrow][mycol]);

			}
			FileOutputStream fOut = new FileOutputStream(outFile); 
			wb.write(fOut);
			fOut.flush();
			fOut.close();
		}


		wb = null;
		osheet = null;

	}
	// Keyword function 
	public static void openBrowser(String fTD){
		if (fTD.equals("Chrome")){
			System.setProperty("webdriver.chrome.driver","C:\\selenium\\chromedriver.exe");
			//Open Browser	chrome	-
			myD = new ChromeDriver();
		}else if (fTD.equals("edge")){
			//System.setProperty("webdriver.chrome.driver","C:\\selenium\\chromedriver.exe");
			myD = new EdgeDriver();
		}else if (fTD.equals("IE")){
			//System.setProperty("webdriver.chrome.driver","C:\\selenium\\chromedriver.exe");
			myD = new InternetExplorerDriver();
		}else if (fTD.equals("FF")){
			//System.setProperty("webdriver.chrome.driver","C:\\selenium\\chromedriver.exe");
			myD = new FirefoxDriver();
		}
		myD.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		myD.manage().window().maximize();
	}
	public static void navigateBrowser(String fTD){
		//Go to AUT
		//I/P: TD (Browser Name)
		// O/P: Nil
		myD.navigate().to(fTD);
	}
	public  static void typeText(String fTD, String fEID){
		// Type into Text Field 
		// I/P: TD (what to type), EID(where to type)
		// O/p: Nil
		myD.findElement(By.xpath(fEID)).sendKeys(fTD);
	}
	public  static void clickElement( String fEID){
		// Type into Text Field 
		// I/P: TD (what to type), EID(where to type)
		// O/p: Nil
		myD.findElement(By.xpath(fEID)).click();
	} 

	public  static void verifyText(String fTD, String fEID){
		// Verify Text in an element 
		// I/P: TD (what to type), EID(where to type)
		// O/p: Nil
		String fActualText;
		fActualText= myD.findElement(By.xpath(fEID)).getText();
		System.out.println("Actual Text is "+ fActualText);
		System.out.println("Expected Text is "+ fTD);


		if (fActualText.equals(fTD)){
			System.out.println("Test matched");
		}else {
			System.out.println("Text did not match");
		}
	} 
	//closeBrowser
	public  static void closeBrowser( ){
		myD.close();

	}

}
