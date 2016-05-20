package com.tries.mergeexcel.xlstoxlsx;

import java.io.IOException;

/**
 * Hello world!
 *
 */
public class App 
{
	private  static final String FILE_TYPE_XLSX=".xlsx";
	
    public static void main( String[] args ) throws IOException
    {
    	
	    ExcelOperation exOperation=new ExcelOperation("C:\\Users\\sozl657\\Desktop\\Results\\merge"+FILE_TYPE_XLSX);
	    exOperation.readXlsFile();	   
	    exOperation.writeToXlsxFile();
    	
    	
        
    }
}
