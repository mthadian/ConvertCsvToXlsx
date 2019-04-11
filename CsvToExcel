import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.lang3.SystemUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;

public class CsvToExcel
{
		 
	    public static final char FILE_DELIMITER = ',';
	    public static final String FILE_EXTN = ".xlsx";
	    public static final String FILE_NAME = "EXCEL_DATA";
	 
	   // private static Logger logger = Logger.getLogger(CsvToExcel.class);
	    private static final Logger logger = LoggerFactory.getLogger(CsvToExcel.class);
	 
	    public static String convertCsvToXls() 
	    {
	    	String slash="";
			
			if (SystemUtils.IS_OS_WINDOWS)
			{
				slash="\\";
			}
			else 
			{
				slash="/";
			}
			
	        SXSSFSheet sheet = null;
	        CSVReaderBuilder  builder = null;
	        CSVReader reader =null;
	        
	        Workbook workBook = null;
	        String generatedXlsFilePath = "";
	        FileOutputStream fileOutputStream = null;
	        String currentWorkingDir = System.getProperty("user.dir");
			String inputFolder=currentWorkingDir.concat(slash+"input");		
			String outputFolder=currentWorkingDir.concat(slash+"output");
			String errorFolder=currentWorkingDir.concat(slash+"error");
			String backUpFolder=currentWorkingDir.concat(slash+"backup");
			
			DateFormat dateFormat = new SimpleDateFormat("yyyy_MM_dd_HH_mm_ss");
			Date gdate = new Date();
			String dateNow=dateFormat.format(gdate);
			
			String newline = System.lineSeparator();

			
			File folderInput= new File(inputFolder);
			File[] files=folderInput.listFiles();
			
			for(File file:files)
			{
				String currentFile=file.getName();
				if(currentFile.contains(".csv"))
				{
					String fileName=currentFile.substring(0, currentFile.length()-4);
					String csvFilePath=inputFolder.concat(slash+currentFile);
			        String  xlsFileLocation=inputFolder.concat(slash+fileName+".xlsx");
			        try {
			        	logger.info("FILE PROCESSING IS "+currentFile,newline);
			       	 
			            /**** Get the CSVReader Instance & Specify The Delimiter To Be Used ****/
			            String[] nextLine;
			            builder = new CSVReaderBuilder(new FileReader(csvFilePath));
			 
			            workBook = new SXSSFWorkbook();
			            sheet = (SXSSFSheet) workBook.createSheet(fileName);
			 
			            int rowNum = 0;
			           // logger.info("Creating New .Xls File From The Already Generated .Csv File");
			            reader = builder.build();
			            while((nextLine = reader.readNext()) != null) {
			                Row currentRow = sheet.createRow(rowNum++);
			                for(int i=0; i < nextLine.length; i++)
			                {
			                	
			                    if(NumberUtils.isDigits(nextLine[i])) 
			                    { 
			                        currentRow.createCell(i).setCellValue(Integer.parseInt(nextLine[i]));
			                    } else if (NumberUtils.isCreatable(nextLine[i])) 
			                    {
			                        currentRow.createCell(i).setCellValue(Double.parseDouble(nextLine[i]));
			                    } else {
			                        currentRow.createCell(i).setCellValue(nextLine[i]);
			                    }
			                }
			            }
			 
			           // generatedXlsFilePath = xlsFileLocation + FILE_NAME + FILE_EXTN;
			            generatedXlsFilePath = xlsFileLocation;
			           // logger.info("The File Is Generated At The Following Location?= " + generatedXlsFilePath);
			 
			            fileOutputStream = new FileOutputStream(generatedXlsFilePath.trim());
			            workBook.write(fileOutputStream);
			            file.delete();
			        } catch(Exception exObj) 
			        {
			        	logger.error(exObj.getMessage(),newline);
			        	exObj.printStackTrace();
			        	String error_currentFile=currentFile.substring(0, 0) + dateNow +" "+currentFile.substring(0);
						file.renameTo(new File(errorFolder+slash+error_currentFile));
			        	
			  
			        } finally
			        {         
			            try
			            {
			 
			                /**** Closing The Excel Workbook Object ****/
			                workBook.close();
			 
			                /**** Closing The File-Writer Object ****/
			                fileOutputStream.close();
			 
			                /**** Closing The CSV File-ReaderObject ****/
			              reader.close();
			              file.delete();
			            } catch (IOException ioExObj) {
			               // logger.error("Exception While Closing I/O Objects In convertCsvToXls() Method?=  " + ioExObj);          
			            }
			        }
					
					
				}
				
				
			}
			
	       // String csvFilePath=inputFolder.concat(slash+"02.04.19 SAFARICOM AIRTIME.csv");
	       // String  xlsFileLocation=inputFolder.concat(slash+"02.04.19 SAFARICOM AIRTIME.xlsx");
	        
	 
	       
	 
	        return generatedXlsFilePath;
	    }   
	

}
