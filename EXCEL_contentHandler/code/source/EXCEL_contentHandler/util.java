package EXCEL_contentHandler;

// -----( IS Java Code Template v1.2
// -----( CREATED: 2004-10-26 16:44:05 JST
// -----( ON-HOST: xiandros-c640

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import java.text.SimpleDateFormat;
import java.util.GregorianCalendar;
// --- <<IS-END-IMPORTS>> ---

public final class util

{
	// ---( internal utility methods )---

	final static util _instance = new util();

	static util _newInstance() { return new util(); }

	static util _cast(Object o) { return (util)o; }

	// ---( server methods )---




	public static final void MSExcelDocumentToRecord (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(MSExcelDocumentToRecord)>> ---
		// @sigtype java 3.5
		// [i] field:0:optional format {"FREEFORM","REPORT"}
		// [o] record:1:required nodelist
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		
			String	file_name = "";
			BufferedInputStream	file_stream = null;
			String	file_data = "";
		
			// data
			IData	data = IDataUtil.getIData( pipelineCursor, "data" );
			if ( data != null)
			{
				IDataCursor dataCursor = data.getCursor();
					file_name = IDataUtil.getString( dataCursor, "file_name" );
					file_stream = (BufferedInputStream)IDataUtil.get( dataCursor, "file_stream" );
					file_data = IDataUtil.getString( dataCursor, "file_data" );
				dataCursor.destroy();
			}
			String	format = IDataUtil.getString( pipelineCursor, "format" );
		pipelineCursor.destroy();
		
		// pipeline
		IDataCursor pipelineCursor_1 = pipeline.getCursor();
		
		IData[]	row_list = null;
		try
		{
			HSSFWorkbook wb = null;
			if (file_name != null)
		    	wb = new HSSFWorkbook(new FileInputStream(file_name));
			else if (file_stream != null)
				wb = new HSSFWorkbook(file_stream);
			else if (file_data != null)
				wb = new HSSFWorkbook(new ByteArrayInputStream(file_data.getBytes()));
		
		    HSSFSheet sheet = wb.getSheetAt(0);
		    HSSFRow row = null;
		    HSSFCell cell = null;
		    String cl = "";
		    double icl = 0;
		
			//////////////////////////////////////////////////////////////
			// Allocate and build field list from spreadsheet
			int cell_count = sheet.getRow(0).getPhysicalNumberOfCells();
			String[] field_name = new String[cell_count];
		
			// Set up field names based on options
			row = sheet.getRow(0);
		
		    for (int k = 0; k < 21; k++)
		    {
		
				
		         cell = row.getCell((short)k);
				
			
		         if (cell != null && format != null)
		         {
					if (format.equals("REPORT"))
					{
		            	if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING){
									
		                		field_name[k] = build_name(cell.getStringCellValue());
								
							}
						else {
							field_name[k] = "C"+Integer.toString(k);	
							
						}
					}
					else field_name[k] = "C"+Integer.toString(k);
				 }
				 else field_name[k] = "C"+Integer.toString(k);
		    }
		
			//////////////////////////////////////////////////////////////
			// Read Excel data and create dynamic record based on fileds
		
			int	row_cnt = sheet.getPhysicalNumberOfRows();
			
			row_list = new IData[row_cnt];
		
			for (int i = 0; i < row_cnt; i++)
			{
				row = sheet.getRow(i);
				IDataCursor idc_row_node = null;
				row_list[i] = IDataFactory.create();
				idc_row_node = row_list[i].getCursor();
		
		    	for (int j = 0; j < row.getPhysicalNumberOfCells(); j++)
		    	{
					cell = row.getCell((short)j);
		        	if (cell != null)
		        	{
		            	int type = cell.getCellType();
		
							switch (type) 
							{
		              			case HSSFCell.CELL_TYPE_STRING:
		 	               			cl = cell.getStringCellValue();
									
										if(field_name[j].equals("Currency")){
											field_name[j]="CurrencyOne";
										}
										if(field_name[j].equals("Coin")){
											field_name[j]="CoinPenny";
										}
										if(field_name[j].equals("Bundles")){
											field_name[j]="BundleTwenty";
										}
									
									if (format.equals("REPORT"))
		                				IDataUtil.put( idc_row_node, field_name[j], cl );
		
												
									else
										IDataUtil.put( idc_row_node, "C"+Integer.toString(j), cl );
		              			break;
		           	
		            			case HSSFCell.CELL_TYPE_NUMERIC:
		                			icl = cell.getNumericCellValue();
		                    		if (isCellDateFormatted(cell))
		                    		{
		                        		// format in form of M/D/YY
		                        		Calendar cal = Calendar.getInstance();
		                        		cal.setTime(getJavaDate(icl,false));
		                        		String pattern = "dd-MMM-yy";
		                        		SimpleDateFormat df = new SimpleDateFormat(pattern);
		                        		String dateStr = df.format(cal.getTime());
										if (format.equals("REPORT"))
											IDataUtil.put( idc_row_node, field_name[j], dateStr);
										else
											IDataUtil.put( idc_row_node, "C"+Integer.toString(j), dateStr );
		
		                    		}
		                    		else
		                				IDataUtil.put( idc_row_node, field_name[j], Double.toString(icl));
								break;
		            			case HSSFCell.CELL_TYPE_BLANK:
			                		icl = cell.getNumericCellValue();
									
									if (format.equals("REPORT")){
										
										if(field_name[j].equals("C5")){
											field_name[j]="CurrencyTwo";
										}
										if(field_name[j].equals("C6")){
											field_name[j]="CurrencyFive";
										}
										if(field_name[j].equals("C7")){
											field_name[j]="CurrencyTen";
										}
										if(field_name[j].equals("C8")){
											field_name[j]="CurrencyTwenty";
										}
										if(field_name[j].equals("C9")){
											field_name[j]="CurrencyFifty";
										}
										if(field_name[j].equals("C10")){
											field_name[j]="CurrencyHundred";
										}
										if(field_name[j].equals("C12")){
											field_name[j]="CoinNickle";
										}
										if(field_name[j].equals("C13")){
											field_name[j]="CoinDime";
										}
										if(field_name[j].equals("C14")){
											field_name[j]="CoinQuarter";
										}
										if(field_name[j].equals("C15")){
											field_name[j]="CoinHalf";
										}
										if(field_name[j].equals("C16")){
											field_name[j]="CoinDollar";
										}
										if(field_name[j].equals("C18")){
											field_name[j]="BundleTen";
										}
										if(field_name[j].equals("C19")){
											field_name[j]="BundleFive";
										}
										if(field_name[j].equals("C20")){
											field_name[j]="BundleOne";
										}
									IDataUtil.put( idc_row_node, field_name[j], ""); }
									else
										
										IDataUtil.put( idc_row_node, "C"+Integer.toString(j), "" );
								break;
		            		} // END SWITCH
				 	}
					else
					{
									if (format.equals("REPORT"))
											IDataUtil.put( idc_row_node, field_name[j], "");
									else
										IDataUtil.put( idc_row_node, "C"+Integer.toString(j), "" );
										
					}
		    	} // End of J FOR
			}
			// Adjust array for REPORT - remove first row
			if (format != null && format.equals("REPORT"))
			{
				IData[]	temp_row = new IData[row_list.length-2];
				for (int i = 0; i < temp_row.length; i++)
				temp_row[i] = row_list[i+2];
				row_list = temp_row;
				temp_row = null;
			}
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
		IDataUtil.put( pipelineCursor_1, "nodelist", row_list );
		pipelineCursor_1.destroy();
		
		// pipeline
		IDataCursor pipelineCursor_2 = pipeline.getCursor();
		
		// recordMSExcel
		IData	recordMSExcel = IDataFactory.create();
		IDataCursor recordMSExcelCursor = recordMSExcel.getCursor();
		
		IDataUtil.put( recordMSExcelCursor, "recordMSExcel", row_list );
		recordMSExcelCursor.destroy();
		
		IDataUtil.put( pipelineCursor, "recordMSExcel", recordMSExcel );
		
		pipelineCursor.destroy();         
		
		
		// --- <<IS-END>> ---

                
	}



	public static final void MSExcelWorkSheetToRecord (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(MSExcelWorkSheetToRecord)>> ---
		// @sigtype java 3.5
		// [i] field:0:optional format {"FREEFORM","REPORT"}
		// [o] record:1:required nodelist
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		
			String	file_name = "";
			BufferedInputStream	file_stream = null;
			String	file_data = "";
		
			// data
			IData	data = IDataUtil.getIData( pipelineCursor, "data" );
			if ( data != null)
			{
				IDataCursor dataCursor = data.getCursor();
					file_name = IDataUtil.getString( dataCursor, "file_name" );
					file_stream = (BufferedInputStream)IDataUtil.get( dataCursor, "file_stream" );
					file_data = IDataUtil.getString( dataCursor, "file_data" );
				dataCursor.destroy();
			}
			String	format = IDataUtil.getString( pipelineCursor, "format" );
		pipelineCursor.destroy();
		
		// pipeline
		IDataCursor pipelineCursor_1 = pipeline.getCursor();
		IData[] work_sheet_list = null;
		IData[]	row_list = null;
		try
		{
			HSSFWorkbook wb = null;
			if (file_name != null)
		    	wb = new HSSFWorkbook(new FileInputStream(file_name));
			else if (file_stream != null)
				wb = new HSSFWorkbook(file_stream);
			else if (file_data != null)
				wb = new HSSFWorkbook(new ByteArrayInputStream(file_data.getBytes()));
		
		    HSSFSheet sheet = null;
		    HSSFRow row = null;
		    HSSFCell cell = null;
		    String cl = "";
		    double icl = 0;
		
			work_sheet_list = new IData[wb.getNumberOfSheets()];
			IDataCursor	idc_sheet_node = null;
		for (int ws = 0; ws < wb.getNumberOfSheets(); ws++)
		{
			sheet = wb.getSheetAt(ws);
			work_sheet_list[ws] = IDataFactory.create();
			idc_sheet_node = work_sheet_list[ws].getCursor();
		
			//////////////////////////////////////////////////////////////
			// Allocate and build field list from spreadsheet
			int cell_count = sheet.getRow(0).getPhysicalNumberOfCells();
			String[] field_name = new String[cell_count];
		
			// Set up field names based on options
			row = sheet.getRow(0);
			
		    for (int k = 0; k < cell_count; k++)
		    {
		         cell = row.getCell((short)k);
		         if (cell != null && format != null)
		         {
					if (format.equals("REPORT"))
					{
		            	if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING)
		                	field_name[k] = build_name(cell.getStringCellValue());
						else field_name[k] = "C"+Integer.toString(k);	
					}
					else field_name[k] = "C"+Integer.toString(k);
				 }
				 else field_name[k] = "C"+Integer.toString(k);
		    }
		
			//////////////////////////////////////////////////////////////
			// Read Excel data and create dynamic record based on fileds
		
			int	row_cnt = sheet.getPhysicalNumberOfRows();
			row_list = new IData[row_cnt];
			
			for (int i = 0; i < row_cnt; i++)
			{
				row = sheet.getRow(i);
				IDataCursor idc_row_node = null;
				row_list[i] = IDataFactory.create();
				idc_row_node = row_list[i].getCursor();
		
		    	for (int j = 0; j < row.getPhysicalNumberOfCells(); j++)
		    	{
		        	cell = row.getCell((short)j);
		        	if (cell != null)
		        	{
		            	int type = cell.getCellType();
		
							switch (type) 
							{
		              			case HSSFCell.CELL_TYPE_STRING:
		 	               			cl = cell.getStringCellValue();
									if (format.equals("REPORT"))
		                				IDataUtil.put( idc_row_node, field_name[j], cl );
									else
										IDataUtil.put( idc_row_node, "C"+Integer.toString(j), cl );
		              			break;
		           	
		            			case HSSFCell.CELL_TYPE_NUMERIC:
		                			icl = cell.getNumericCellValue();
		                    		if (isCellDateFormatted(cell))
		                    		{
		                        		// format in form of M/D/YY
		                        		Calendar cal = Calendar.getInstance();
		                        		cal.setTime(getJavaDate(icl,false));
		                        		String pattern = "dd-MMM-yy";
		                        		SimpleDateFormat df = new SimpleDateFormat(pattern);
		                        		String dateStr = df.format(cal.getTime());
										if (format.equals("REPORT"))
											IDataUtil.put( idc_row_node, field_name[j], dateStr);
										else
											IDataUtil.put( idc_row_node, "C"+Integer.toString(j), dateStr );
		                    		}
		                    		else
		                				IDataUtil.put( idc_row_node, field_name[j], Double.toString(icl));
								break;
		            			case HSSFCell.CELL_TYPE_BLANK:
			                		icl = cell.getNumericCellValue();
									if (format.equals("REPORT"))
		    	            			IDataUtil.put( idc_row_node, field_name[j], "");
									else
										IDataUtil.put( idc_row_node, "C"+Integer.toString(j), "" );
								break;
		            		} // END SWITCH
				 	}
					else
					{
						if (format.equals("REPORT"))
		           			IDataUtil.put( idc_row_node, field_name[j], "");
						else
							IDataUtil.put( idc_row_node, "C"+Integer.toString(j), "" );
					}
		    	} // End of J FOR
			}
			// Adjust array for REPORT - remove forst row
			if (format != null && format.equals("REPORT"))
			{
				IData[]	temp_row = new IData[row_list.length-1];
				for (int i = 0; i < temp_row.length; i++)
					temp_row[i] = row_list[i+1];
				row_list = temp_row;
				temp_row = null;
			}
			// Setup worksheet record
			IDataUtil.put( idc_sheet_node, "row", row_list);
		
		}
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
		
		pipelineCursor_1.destroy();
		
		// pipeline
		IDataCursor pipelineCursor_2 = pipeline.getCursor();
		
		// recordMSExcel
		IData	recordMSExcel = IDataFactory.create();
		IDataCursor recordMSExcelCursor = recordMSExcel.getCursor();
		
		IDataUtil.put( recordMSExcelCursor, "recordMSExcel", work_sheet_list );
		recordMSExcelCursor.destroy();
		
		IDataUtil.put( pipelineCursor, "recordMSExcel", recordMSExcel );
		
		pipelineCursor.destroy();
		// --- <<IS-END>> ---

                
	}



	public static final void RecordToMSExcel (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(RecordToMSExcel)>> ---
		// @subtype unknown
		// @sigtype java 3.5
		// [i] field:0:required file
		// [i] record:1:required in_doc
		
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
			String	file = IDataUtil.getString( pipelineCursor, "file" );
		
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet("new sheet");
		HSSFRow row = null;
		HSSFCell cell = null;
		
		try {
		
		    String key = "";
		    String val = "";
		    Object valObj = null;
		/*
		    // Create a cell.
		    row.createCell((short)0).setCellValue(1);
		    row.createCell((short)1).setCellValue(1.2);
		    row.createCell((short)2).setCellValue("This is a string");
		    row.createCell((short)3).setCellValue(true);
		*/
		    // in_doc
		    IData[] in_doc = IDataUtil.getIDataArray( pipelineCursor, "in_doc" );
		    if ( in_doc != null)
		    {
			// Handle all records - rows
			for ( int i = 0; i < in_doc.length; i++ )
			{
		    	     // Create a row and put some cells in it. Rows are 0 based.
		    	     row = sheet.createRow((short)i);
			     
			     int count = 0;
			     IDataCursor idc = in_doc[i].getCursor();
			     idc.first();
			     boolean more_data = true;
			     while (more_data)//(idc.hasMoreData())
			     {
				key = idc.getKey();
				val = (String)idc.getValue();
		    
				// Create a cell.
				row.createCell((short)count).setCellValue(val);
				count++;
		
			        more_data = idc.next();
			     }
		             idc.destroy();
			}
		    }
		    pipelineCursor.destroy();
		
		    // Write the output to a file
		    FileOutputStream fileOut = new FileOutputStream(file);
		    wb.write(fileOut);
		    fileOut.close();
		
		} catch (Exception e) {
			//throw new ServiceException(e.getMessage());
			e.printStackTrace();
		}
		// --- <<IS-END>> ---

                
	}



	public static final void createExcelRecord (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(createExcelRecord)>> ---
		// @sigtype java 3.5
		// [i] field:0:optional format {"REPORT","FREEFORM"}
		// [o] record:0:required boundnode
		
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		
			String	file_name = "";
			BufferedInputStream	file_stream = null;
			String	file_data = "";
		
			// data
			IData	data = IDataUtil.getIData( pipelineCursor, "data" );
			if ( data != null)
			{
				IDataCursor dataCursor = data.getCursor();
					file_name = IDataUtil.getString( dataCursor, "file_name" );
					file_stream = (BufferedInputStream)IDataUtil.get( dataCursor, "file_stream" );
					file_data = IDataUtil.getString( dataCursor, "file_data" );
				dataCursor.destroy();
			}
			String	format = IDataUtil.getString( pipelineCursor, "format" );
		pipelineCursor.destroy();
		
		// pipeline
		IDataCursor pipelineCursor_1 = pipeline.getCursor();
		// boundnode
		IData	boundnode = IDataFactory.create();
		// pipeline
		IDataCursor pipeline_boundnode = boundnode.getCursor();
		try
		{
			HSSFWorkbook wb = null;
			if (file_name != null)
		    	wb = new HSSFWorkbook(new FileInputStream(file_name));
			else if (file_stream != null)
				wb = new HSSFWorkbook(file_stream);
			else if (file_data != null)
				wb = new HSSFWorkbook(new ByteArrayInputStream(file_data.getBytes()));
		
		    HSSFSheet sheet = wb.getSheetAt(0);
		    HSSFRow row = sheet.getRow(0);
		    HSSFCell cell = null;
		    String cl = "";
		    double icl = 0;
		
		      for (int j = 0; j < row.getPhysicalNumberOfCells(); j++)
		      {
		         cell = row.getCell((short)j);
		         if (cell != null && format != null)
		         {
					if (format.equals("REPORT"))
					{
		            	int type = cell.getCellType();
		            	if (type == HSSFCell.CELL_TYPE_STRING)
		            	{
		                	cl = cell.getStringCellValue();
		                	//System.out.println("Row - " + Integer.toString(i) + " Cell - " + Integer.toString(j) + " "+build_name(cl));
							IDataUtil.put( pipeline_boundnode, build_name(cl), "" );
		            	}
					}
					else IDataUtil.put( pipeline_boundnode, "C"+Integer.toString(j), "" );
				 }
				 else  IDataUtil.put( pipeline_boundnode, "C"+Integer.toString(j), "" );
		      }
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
		IDataUtil.put( pipelineCursor_1, "boundnode", boundnode );
		pipelineCursor_1.destroy();
		// --- <<IS-END>> ---

                
	}

	// --- <<IS-START-SHARED>> ---
	
	private static final long   DAY_MILLISECONDS  = 24 * 60 * 60 * 1000;
	
	public static String build_name(String str)
	{
	  String name = "";
	  StringTokenizer strtok = new StringTokenizer(str," ");
	  String temp = "";
	  int count = 0;
	  while (strtok.hasMoreElements())
	  {
	    temp = (String)strtok.nextElement();
	    if (count == 0)
	        name = temp;
	    else name += "_"+temp;
	
	    count++;
	  }
	  return name;
	}
	      /**
	       * Given a double, checks if it is a valid Excel date.
	       *
	       * @return true if valid
	       * @param  value the double value
	       */
	      public static boolean isValidExcelDate(double value)
	      {
	          return (value > -Double.MIN_VALUE);
	      }
	
	//////////////////////////////////////////////////////////////////
	// method to determine if the cell is a date, versus a number...
	public static boolean isCellDateFormatted(HSSFCell cell) 
	{
	    boolean bDate = false;
	
	    double d = cell.getNumericCellValue();
	    if ( isValidExcelDate(d) ) {
	      HSSFCellStyle style = cell.getCellStyle();
	      int i = style.getDataFormat();
	      switch(i) {
	    // Internal Date Formats as described on page 427 in Microsoft Excel Dev's Kit...
	        case 0x0e:
	        case 0x0f:
	        case 0x10:
	        case 0x11:
	        case 0x12:
	        case 0x13:
	        case 0x14:
	        case 0x15:
	        case 0x16:
	        case 0x2d:
	        case 0x2e:
	        case 0x2f:
	         bDate = true;
	        break;
	
	        default:
	         bDate = false;
	        break;
	      }
	    }
	    return bDate;
	  }
	
	      /**
	       * Given a Calendar, return the number of days since 1600/12/31.
	       *
	       * @return days number of days since 1600/12/31
	       * @param  cal the Calendar
	       * @exception IllegalArgumentException if date is invalid
	       */
	
	      private static int absoluteDay(Calendar cal)
	      {
	          return cal.get(Calendar.DAY_OF_YEAR)
	                 + daysInPriorYears(cal.get(Calendar.YEAR));
	      }
	
	      /**
	       * Return the number of days in prior years since 1601
	       *
	       * @return    days  number of days in years prior to yr.
	       * @param     yr    a year (1600 < yr < 4000)
	       * @exception IllegalArgumentException if year is outside of range.
	       */
	
	      private static int daysInPriorYears(int yr)
	      {
	          if (yr < 1601)
	          {
	              throw new IllegalArgumentException(
	                  "'year' must be 1601 or greater");
	          }
	          int y    = yr - 1601;
	          int days = 365 * y      // days in prior years
	                     + y / 4      // plus julian leap days in prior years
	                     - y / 100    // minus prior century years
	                     + y / 400;   // plus years divisible by 400
	
	          return days;
	      }
	
	      /**
	       *  Given an Excel date with either 1900 or 1904 date windowing,
	       *  converts it to a java.util.Date.
	       *
	       *  @param date  The Excel date.
	       *  @param use1904windowing  true if date uses 1904 windowing,
	       *   or false if using 1900 date windowing.
	       *  @return Java representation of the date, or null if date is not a valid Excel date
	       */
	      public static Date getJavaDate(double date, boolean use1904windowing) {
	          if (isValidExcelDate(date)) {
	              int startYear = 1900;
	              int dayAdjust = -1; // Excel thinks 2/29/1900 is a valid date, which it isn't
	              int wholeDays = (int)Math.floor(date);
	              if (use1904windowing) {
	                  startYear = 1904;
	                  dayAdjust = 1; // 1904 date windowing uses 1/2/1904 as the first day
	              }
	              else if (wholeDays < 61) {
	                  // Date is prior to 3/1/1900, so adjust because Excel thinks 2/29/1900 exists
	                  // If Excel date == 2/29/1900, will become 3/1/1900 in Java representation
	                  dayAdjust = 0;
	              }
	              GregorianCalendar calendar = new GregorianCalendar(startYear,0, wholeDays + dayAdjust);
	              int millisecondsInDay = (int)((date - Math.floor(date)) * (double) DAY_MILLISECONDS + 0.5);
	              calendar.set(GregorianCalendar.MILLISECOND, millisecondsInDay);
	              return calendar.getTime();
	          }
	          else {
	              return null;
	          }
	      }
	// --- <<IS-END-SHARED>> ---
}

