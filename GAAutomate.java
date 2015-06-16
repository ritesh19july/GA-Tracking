package com.selenium.performancetest;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.io.*;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import org.browsermob.core.har.Har;
import org.browsermob.proxy.ProxyServer;
import org.openqa.selenium.By;
import org.openqa.selenium.Proxy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.codehaus.jackson.JsonParseException;

import edu.umass.cs.benchlab.har.*;
import edu.umass.cs.benchlab.har.tools.*;

public class GAAutomate

{
	private String inputFile;

	public static String HARFILE_NAME="PerformanceTestHar.har";
	public static String TEXTFILE_NAME="PerformanceTestHar.txt";
	public static String EXCELFILE_NAME="GA_Excel_Execution.xls";
	
	public void setInputFile(String inputFile) 
	{
		this.inputFile = inputFile;
	}
	
	public String readCellDataUsingIndexFromXLSFile(String inputFile,int workSheetNumber, int ColIndex, int rowIndex) throws IOException {
		String cellData = null;

		File inputWorkbook = new File(inputFile);
		Workbook w;

		try {
			w = Workbook.getWorkbook(inputWorkbook);

			// Get the first sheet
			Sheet sheet = w.getSheet(workSheetNumber);
			System.out.println("Specific Row data: " + sheet);

			WorkbookSettings ws = new WorkbookSettings();
			ws.setLocale(new Locale("er", "ER"));

			//Total No. of Columns
			int columns = sheet.getColumns();
			System.out.println("No. of Columns : " + columns);

			//Total No. of Rows
			int rows = sheet.getRows();
			System.out.println("No. of Rows : " + rows);

			// Loop over first 10 column and lines
			int i = 0, j = 0;
			for (i = 1; i < rows; i++) {
				if (i == rowIndex) 
				{
					break;
				}
			}
			for (j = 0; j < columns; j++) {
				if (ColIndex == j) {
					break;
				}
			}
			cellData = sheet.getCell(j, i).getContents();
			System.out.println("Cell Data :: " + cellData);

		} catch (BiffException e) 
		{
			e.printStackTrace();
		}
		
		System.out.println("Specific Row data: " + cellData);

		;

		return cellData;
	}
	public void writeCellDataInExistingXLS(String excelPath, int SheetIndex,int rowIndex, int ColIndex, String Text) throws IOException,RowsExceededException, WriteException, BiffException 
	
	{
		Workbook aWorkBook = Workbook.getWorkbook(new File(excelPath));
		WritableWorkbook aCopy = Workbook.createWorkbook(new File(excelPath),aWorkBook);
		WritableSheet aCopySheet = aCopy.getSheet(SheetIndex);// index of the needed sheet
		jxl.write.Label anotherWritableCell = new jxl.write.Label(ColIndex,rowIndex, Text);
		System.out.print(ColIndex);
		
		aCopySheet.addCell(anotherWritableCell);
		aCopy.write();
		aCopy.close();
	}

	public void writeLogToTextFile()
	{
		try{
			
			String workingDir = System.getProperty("user.dir");
			String filename = workingDir+"\\"+GAAutomate.HARFILE_NAME;
			String filename1 = workingDir+"\\"+GAAutomate.TEXTFILE_NAME;
			
			File f = new File(filename);
			HarFileReader readhar = new HarFileReader();
			HarFileWriter writehar = new HarFileWriter();
			HarLog log = readhar.readHarFile(f);
			// Access all elements as objects
			HarBrowser browser = log.getBrowser();
			HarEntries entries = log.getEntries();
			// Used for loops
			List<HarPage> pages = log.getPages().getPages();
			List<HarEntry> hentry = entries.getEntries();
			for (HarPage page : pages) {
				System.out.println("page start time: " + ISO8601DateFormatter.format(page.getStartedDateTime()));
				System.out.println("page id: " + page.getId());
				System.out.println("page title: " + page.getTitle());
			}
			// Output "response" code of entries.
			for (HarEntry entry : hentry) 
				{
				System.out.println("request code:  "+ entry.getRequest().getMethod()); // Output request// type
				System.out.println("response code: "+ entry.getRequest().getUrl()); // Output url of request
				System.out.println("response code: "+ entry.getResponse().getStatus()); // Output the
			}

			// Once you are done manipulating the objects, write back to a file
			System.out.println("Writing " + EXCELFILE_NAME );
			
			File f2 = new File(filename1);
			writehar.writeHarFile(log, f2);
			
		}catch(Exception e)
		{
			
		}
	}
	
	public ArrayList<String> readUtmeFromTextFile()
	{
		ArrayList<String> utmeList=new ArrayList<String>();
		String workingDir = System.getProperty("user.dir");
		String filename1 = workingDir+"\\"+GAAutomate.TEXTFILE_NAME;
		
		try{
			
			String line = "";
			BufferedReader br = new BufferedReader(new FileReader(filename1));
			while ((line = br.readLine()) != null) 
			{
				if (line.contains("\"url\"")) 
				{
					if (line.indexOf("utme") != -1) 
					{
						utmeList.add(line.substring(line.indexOf("utme"),line.indexOf("&utmcs")));
						
					}
				}

			}
			
		}catch(Exception e){
			
		}
		
		return utmeList;
	}
		
	
	public ArrayList<String> readCellData(String inputFile,int ColIndex) throws IOException {
		String cellData = null;

		File inputWorkbook = new File(inputFile);
		Workbook w;
		ArrayList<String> urlList=new ArrayList<String>();
		
		try {
			w = Workbook.getWorkbook(inputWorkbook);
			// Get the first sheet
			Sheet sheet = w.getSheet("Sheet1");
			System.out.println("Specific Row data: " + sheet);
			WorkbookSettings ws = new WorkbookSettings();
			ws.setLocale(new Locale("er", "ER"));
			int columns = sheet.getColumns();
			int rows = sheet.getRows();
			

			// Loop over first 10 column and lines
			
			for (int i = 1; i < sheet.getRows(); i++) 
			{
				cellData = sheet.getCell(ColIndex, i).getContents();
				System.out.println(cellData);
				urlList.add(cellData);
			}
			
			
			System.out.println("Cell Data :: " +urlList);

		} catch (Exception e) 
		{
			e.printStackTrace();
		}
		
		

		;

		return urlList;
	}
	
	public static void main(String[] args)throws Exception 
	{
		
		String workingDir = System.getProperty("user.dir");
		System.out.println("Current working directory : " + workingDir);
		String filename = workingDir+"\\"+GAAutomate.HARFILE_NAME;
		GAAutomate read = new GAAutomate();
		
		
		
		
		ArrayList<String> urlList=read.readCellData(workingDir+ "\\"+EXCELFILE_NAME, 0);
		String cellData = null;

		File inputWorkbook = new File(EXCELFILE_NAME);
		Workbook w;
		
		
		
			w = Workbook.getWorkbook(inputWorkbook);
			// Get the first sheet
			Sheet sheet = w.getSheet("Sheet1");
			System.out.println("Specific Row data: " + sheet);
			WorkbookSettings ws = new WorkbookSettings();
			ws.setLocale(new Locale("er", "ER"));
			int columns = sheet.getColumns();
			int rows = sheet.getRows();
			

			// Loop over first 10 column and lines
			
			for (int i = 1; i < sheet.getRows(); i++) 
			{
				
		
		
		//System.out.println(ReadUrl);
		String url = read.readCellDataUsingIndexFromXLSFile(workingDir+ "\\"+EXCELFILE_NAME, 0, 0, i);
		
		File f = new File(filename);
		
		//*************************BrowserMob Proxy Started*********************************//
		String PROXY = "localhost:9039";
		
			// start the proxy
	    	ProxyServer server = new ProxyServer(9039);
	    	server.start();
	    
	    	//captures the mouse movements and navigations
	    	server.setCaptureHeaders(true);
        	server.setCaptureContent(true);

	    	// get the Selenium proxy object
	    	Proxy proxy = server.seleniumProxy();

	    	// configure it as a desired capability
	    	DesiredCapabilities capabilities = new DesiredCapabilities();
	    	capabilities.setCapability(CapabilityType.PROXY,proxy);

	    	// start the browser up
	    	WebDriver driver = new InternetExplorerDriver(capabilities);
	    
	    	// create a new HAR with the label "wwe.com"
	    	server.newHar("http://64.152.0.51/");

	    	// open the url
	    	driver.get(url);
	    
	    	// get the HAR data
        	Har har = server.getHar();
        	FileOutputStream fos = new FileOutputStream(filename);
        	har.writeTo(fos);
        	server.stop();
        
        	// Browser Close
		driver.quit();

		//*******************Try Block***********************
		try {
			System.out.println("Reading " + filename);
			
			
			read.writeLogToTextFile();
			ArrayList<String> outputUTME =  read.readUtmeFromTextFile();
			GAAutomate write = new GAAutomate();
			write.writeCellDataInExistingXLS(workingDir+ "\\"+EXCELFILE_NAME,0, i, 1,outputUTME.toString());

			
			
		}
		 catch (Exception e) 
			{
			 e.printStackTrace();   
			} 
		 

	}

	}

}

