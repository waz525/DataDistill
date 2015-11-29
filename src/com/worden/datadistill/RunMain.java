package com.worden.datadistill;

import java.io.*;
import java.util.*;

public class RunMain {
	private int GetMode = 0 ;
	private String Separator = "\t" ;
	private String NewLineSymbol = "\n" ;
	private int OutRowCount = 0 ;
	private int OutCellCount = 0 ;
	private boolean isNeedHeader = false ;
	private int[] OutCellIndexs = null ;
	private String[] OutCellHeaders = null ;
	private boolean isWriteFile = false ;
	private String OutFilePath = "" ;
	private boolean isAddToFile = false ;

	public boolean LoadPropfile( String fileName )
	{

		if ( !FileUtil.isFile(fileName) )
		{
			System.out.println( "[ERROR] "+fileName+" is not a file !!!" ) ;
			return  false;
		}
		Properties prop = new Properties() ;
		try{
			InputStream is = new BufferedInputStream(new FileInputStream(fileName)) ;
			
			prop.load(is) ;
			is.close() ;
		}catch(Exception e )
		{
			e.printStackTrace() ;
		}
		
		String temp ;
		String str[] = null;
		String CharsetConf = "UTF-8" ;
		
		temp = prop.getProperty("GetMode");
		if(  temp != null && temp.length() > 0 ) {
			this.GetMode = Integer.parseInt(temp);
		}

		temp = prop.getProperty("CharsetConf");
		if(  temp != null && temp.length() > 0 ) {
			CharsetConf = temp;
		}

		temp = prop.getProperty("Separator");
		if(  temp != null && temp.length() > 0 ) {
			Separator = temp;
		}
		
		temp = prop.getProperty("NewLineSymbol");
		if(  temp != null && temp.length() > 0 ) {
			NewLineSymbol = temp;
		}
		
		temp = prop.getProperty("OutRowCount");
		if(  temp != null && temp.length() > 0 ) {
			OutRowCount = Integer.parseInt(temp);
		}
		
		
		//获取序列列表
		temp = prop.getProperty("OutCellIndexs");
		if(  temp != null && temp.length() > 0 && GetMode == 0 ) {
			str = temp.split(",") ;

			this.OutCellCount =str.length  ;
			this.OutCellIndexs = new int[this.OutCellCount] ;
			for( int i = 0 ; i<OutCellCount ; i++) {
				this.OutCellIndexs[i] = Integer.parseInt(str[i]);
			}
		}
		
		//获取列头列表
		temp = prop.getProperty("OutCellHeaders");
		if(  temp != null && temp.length() > 0 && GetMode != 0 ) {
			str = temp.split(",") ;

			this.OutCellCount =str.length  ; 
			this.OutCellHeaders = new String[this.OutCellCount] ;
			for( int i = 0 ; i<OutCellCount ; i++) {
				try {
					this.OutCellHeaders[i] = new String(str[i].getBytes(CharsetConf), "utf-8") ;
				} catch (UnsupportedEncodingException e) {
					e.printStackTrace();
				}
			}
		}


		temp = prop.getProperty("isNeedHeader");
		if(  temp != null && temp.length() > 0 ) {
			if( temp.equalsIgnoreCase("true")  ) {
				isNeedHeader = true  ;
			} else {
				isNeedHeader = false ;
			}
		}

		temp = prop.getProperty("isWriteFile");
		if(  temp != null && temp.length() > 0 ) {
			if( temp.equalsIgnoreCase("true")  ) {
				isWriteFile = true  ;
			} else {
				isWriteFile = false ;
			}
		}

		temp = prop.getProperty("OutFilePath");
		if(  temp != null && temp.length() > 0 ) {
			OutFilePath = temp;
		}
		
		
		if( isWriteFile &&  OutFilePath.length() < 1 ) {
			System.err.println("Error: Config is wrong , OutFilePath is needed !!!") ;
			return false ;
		}

		temp = prop.getProperty("isAddToFile");
		if(  temp != null && temp.length() > 0 ) {
			if( temp.equalsIgnoreCase("true")  ) {
				isAddToFile = true  ;
			} else {
				isAddToFile = false ;
			}
		}
		
		return true ;

	}
		
	private void distillExcelFile(String filePath) {
		ExcelReader er = new ExcelReader(filePath) ;
		er.getSheetDataByIndex(0) ;
		int ind = 1  ;
		String ColumnStr = "" ;
		String CellStr = "" ;
		
		if( GetMode == 0 && isNeedHeader ) ind = 0 ;
		for(  ; ind< er.getNuberOfRows() ; ind++ ) {
			ColumnStr = "" ;
			
			for( int i = 0 ; i<this.OutCellCount ; i++) {
				if( GetMode == 0 ) {
					CellStr = er.getCellData(ind, this.OutCellIndexs[i] ) ;
				} else {
					CellStr = er.getCellData(ind, this.OutCellHeaders[i] ) ;
				}
				ColumnStr += CellStr + ( i<this.OutCellCount-1?Separator:NewLineSymbol); 
			}
			
			if( isWriteFile ) {
				FileUtil.WriteString(OutFilePath, ColumnStr);
			} else {
				System.out.print(ColumnStr) ;
			}
			
			
			if( OutRowCount > 0 && ind >= OutRowCount ) break ;
		}		
		
	}
	
	private void distillWordFile(String filePath) {
		WordReader wr = new WordReader(filePath) ;
		wr.getTableContent() ;
		int ind = 1  ;
		String ColumnStr = "" ;
		String CellStr = "" ;
		
		if( GetMode == 0 && isNeedHeader ) ind = 0 ;
		for(  ; ind< wr.getNuberOfRows() ; ind++ ) {
			ColumnStr = "" ;
			for( int i = 0 ; i<this.OutCellCount ; i++) {
				if( GetMode == 0 ) {
					CellStr = wr.getCellData(ind, this.OutCellIndexs[i] ) ;
				} else {
					CellStr = wr.getCellData(ind, this.OutCellHeaders[i] ) ;
				}
				ColumnStr += CellStr + ( i<this.OutCellCount-1?Separator:NewLineSymbol); 
			}
			if( isWriteFile ) {
				FileUtil.WriteString(OutFilePath, ColumnStr);
			} else {
				System.out.print(ColumnStr) ;
			}
			
			if( OutRowCount > 0 && ind > OutRowCount ) break ;
		}		
		
		
	}
	
	public void distillFile( String filePath) {
		String fileType = filePath.substring(filePath.lastIndexOf(".")+1,filePath.length()) ;
		if( isWriteFile ) {
			if( OutFilePath.indexOf("{OldFileName}") >-1 ) {
				String fileName = "" ;
				if( filePath.indexOf("\\") > -1 ) {
					fileName = filePath.substring(filePath.lastIndexOf("\\")+1) ;
				} else if ( filePath.indexOf("/") > -1 ) {
					fileName = filePath.substring(filePath.lastIndexOf("/")+1) ;
				} else {
					fileName = filePath ;
				}
				fileName = fileName.substring(0, fileName.lastIndexOf(".")) ;
				OutFilePath = OutFilePath.replace("{OldFileName}", fileName) ;
			}
			
			if( OutFilePath.indexOf("{DataTime}") >-1 ) {
                java.util.Date nowDate = new java.util.Date() ;
                java.text.SimpleDateFormat formatter = new java.text.SimpleDateFormat("yyyyMMddHHmmss") ;
                String datestr = formatter.format(nowDate);
                OutFilePath = OutFilePath.replace("{DataTime}", datestr) ;
			}

			if( OutFilePath.indexOf("{Data}") >-1 ) {
                java.util.Date nowDate = new java.util.Date() ;
                java.text.SimpleDateFormat formatter = new java.text.SimpleDateFormat("yyyyMMdd") ;
                String datestr = formatter.format(nowDate);
                OutFilePath = OutFilePath.replace("{Data}", datestr) ;
			}

			if( FileUtil.isFile(OutFilePath) && !isAddToFile ) {
				FileUtil.RemoveFile(OutFilePath);
			}
		}
		
		
		if(fileType.equalsIgnoreCase("xls") ||  fileType.equalsIgnoreCase("xlsx") ) {
			distillExcelFile(filePath) ;
		} else if( fileType.equalsIgnoreCase("doc") ||  fileType.equalsIgnoreCase("docx") ) {
			distillWordFile(filePath) ;			
		} else {
			System.out.println("Error: Could not distill "+filePath) ;
		}
			
	}
	
	public static void main(String[] args) {

		if( args.length < 2 )
		{
			System.out.println(" Usage: java com.worden.datadistill.RunMain ConfigFilePath DataFilePath" ) ;
			System.out.println("           Version: 1.0" ) ;
		}
		else
		{
			//System.out.println(args[0] + " --- " + args[1]);
			RunMain process = new RunMain() ;
			if( process.LoadPropfile(args[0]) ) {
				process.distillFile(args[1]);
				System.out.println("\nINFO: Run Over !" ) ;
			}
			
		}
	}

}
