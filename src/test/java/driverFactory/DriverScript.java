package driverFactory;

import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import commonFunction.FunctionLibrary;
import utilities.ExcelFileUtil;

public class DriverScript 
{
	WebDriver driver;
	String inputpath="./FileInput/DataEngine.xlsx";
	String outputpath="./FileOutput/HybridResult.xlsx";
	ExtentReports reports;
	ExtentTest logger;
	//only masterTestCases Sheet
	String TestCases="MasterTestCases";
	public void startTest() throws Throwable
	{
		String module_status="";
		ExcelFileUtil xl= new ExcelFileUtil(inputpath);
		//iterate all test cases in Testcases(MasterTestcases)
		for(int i=1; i<=xl.rowCount(TestCases); i++)
		{
			if(xl.getCellData(TestCases, i, 2).equalsIgnoreCase("Y"))
			{
				//read corresponding sheet or test cases
				String TcModule=xl.getCellData(TestCases, i, 1);
				reports = new ExtentReports("./target/ExtentReports/"+TcModule+FunctionLibrary.generateDate()+".html");
				logger=reports.startTest(TcModule);
				logger.assignAuthor("Sai");
				//Interate all rows in TcModule sheet
				for(int j=1; j<=xl.rowCount(TcModule); j++)
				{
					// read all cells from tcModule sheet
					String Description=xl.getCellData(TcModule, j, 0);
					String Object_type=xl.getCellData(TcModule, j, 1);
					String Lname=xl.getCellData(TcModule, j, 2);
					String Lvalue=xl.getCellData(TcModule, j, 3);
					String Test_data=xl.getCellData(TcModule, j, 4);

					try 
					{
						if(Object_type.equalsIgnoreCase("startBrowser"))
						{
							driver=FunctionLibrary.startBrowser();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("openUrl"))
						{
							FunctionLibrary.openUrl();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("waitForElement"))
						{
							FunctionLibrary.waitForElement(Lname, Lvalue, Test_data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("typeAction"))
						{ 
							FunctionLibrary.typeAction(Lname, Lvalue, Test_data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("clickAction"))
						{
							FunctionLibrary.clickAction(Lname, Lvalue);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("validateTitle"))
						{
							FunctionLibrary.validateTitle(Test_data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("closeBrowser"))
						{
							FunctionLibrary.closeBrowser();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("mouseClick"))
						{
							FunctionLibrary.mouseClick();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("categoryTable"))
						{
							FunctionLibrary.categoryTable(Test_data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("dropDownAction"))
						{
							FunctionLibrary.dropDownAction(Lname, Lvalue, Test_data);
						}
						if(Object_type.equalsIgnoreCase("captureStock"))
						{
							FunctionLibrary.captureStock(Lname, Lvalue);
						}
						if(Object_type.equalsIgnoreCase("stockTable"))
						{
							FunctionLibrary.stockTable();
						}
						if(Object_type.equalsIgnoreCase("captuersupplier"))
						{
							FunctionLibrary.captuersupplier(Lname, Lvalue);
						}
						if(Object_type.equalsIgnoreCase("supplierTable"))
						{
							FunctionLibrary.supplierTable();
						}
						if(Object_type.equalsIgnoreCase("captuerCustomer"))
						{
							FunctionLibrary.captuerCustomer(Lname, Lvalue);
						}
						if(Object_type.equalsIgnoreCase(""))
						{
							FunctionLibrary.customerTable();
						}

						//write as pass into status sheet
						xl.setCellData(TcModule, j, 5, "pass", outputpath);
						logger.log(LogStatus.PASS, Description);
						module_status="True";

					}
					catch(Exception a)
					{
						System.out.println(a.getMessage());
						//write as fail into tc TcModule sheet in status cell
						xl.setCellData(TcModule, j, 3, "fai", outputpath);
						logger.log(LogStatus.FAIL, Description);
						module_status="false";
					}
					if(module_status.equalsIgnoreCase("True"))
					{
						//write as passs into testcases sheet
						xl.setCellData(TestCases, i, 3, "pass", outputpath);
					}
					else
					{
						//write as fail into test cases sheet 
						xl.setCellData(TestCases, i, 3, "fail", outputpath);
					}
					reports.endTest(logger);
					reports.flush();
				}
			}
			else {
				//write as blocked for test cases flat to N in testcases sheet
				xl.setCellData(TestCases, i, 3, "blocked", outputpath);
			}
		}
	}
}
