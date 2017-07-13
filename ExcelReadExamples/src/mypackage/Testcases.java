package mypackage;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.Test;

public class Testcases {
	
	// sheet name will be same for one class/entire functionality of one page
	public String TestfileName="mydata.xls";
	public int sheetNumber;
	@Test
	public void getdataFromexceltest() throws EncryptedDocumentException, FileNotFoundException, InvalidFormatException, IOException
	{
		//sheet number can be decided as per the test scenario/case
		sheetNumber=0;
		
		// getting data in an arraylist row based on search pattern(testcasename=first parameter)
		ArrayList<String> row = Utility.getRow("Test1",TestfileName,sheetNumber);
		
		//not get data based on column number
		String fname=row.get(1);
		System.out.println(fname);
		
		String lname=row.get(2);
		System.out.println(lname);
		
		
		
		//sheet number can be decided as per the test scenario/case
				sheetNumber=1;
				
				// getting data in an arraylist row based on search pattern(testcasename=first parameter)
		row = Utility.getRow("Test24",TestfileName,sheetNumber);
				
				//not get data based on column number
				String taxvalue=row.get(2);
				System.out.println(taxvalue);
				
				//convertit to double if required
			double tax=Double.parseDouble(taxvalue);
			System.out.println("tax="+tax);
		
	}

}
