package com.worden.datadistill;

import java.io.*;
import java.text.DecimalFormat;
import java.util.*;

import org.apache.poi.ss.usermodel.*;

public class ExcelReader {
    private Workbook workBook; 
    private List<String> columnHeaderList = null;
    private List<List<String>> listData;
    private List<Map<String,String>> mapData;
    private boolean flag = false;
    private int numOfRows = 0 ;
    private int numOfColumn = 0 ;
    
    public ExcelReader(String filePath) {
        this.flag = false;
        this.LoadExcelFile(filePath);
    }

    public ExcelReader(String filePath, int index) {
        this.flag = false;
        this.LoadExcelFile(filePath);
        this.getSheetDataByIndex(index);
    } 
    
    public ExcelReader(String filePath, String name) {
        this.flag = false;
        this.LoadExcelFile(filePath);
        this.getSheetDataByName(name);
    }  
    
	/**
	 * 读取Excel文件
	 *
	 */
    private void LoadExcelFile(String filePath) {
        FileInputStream inStream = null;
        try {
            inStream = new FileInputStream(new File(filePath));
            workBook = WorkbookFactory.create(inStream);
        } catch (Exception e) {
            e.printStackTrace();
        }finally{
            try {
                if(inStream!=null){
                    inStream.close();
                }                
            } catch (IOException e) {                
                e.printStackTrace();
            }
        }
    }

	/**
	 * 获取单个单元格内容
	 * @return String
	 */
    private String getCellValue(Cell cell) {
        String cellValue = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        DataFormatter formatter = new DataFormatter();
                        cellValue = formatter.formatCellValue(cell);
                    } else {
                    	DecimalFormat df = new DecimalFormat("0");
                    	cellValue = df.format(cell.getNumericCellValue());
                    	/*
                        double value = cell.getNumericCellValue();
                        int intValue = (int) value;
                        cellValue = value - intValue == 0 ? String.valueOf(intValue) : String.valueOf(value);
                        */
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    cellValue = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    cellValue = String.valueOf(cell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_BLANK:
                    cellValue = "";
                    break;
                case Cell.CELL_TYPE_ERROR:
                    cellValue = "";
                    break;
                default:
                    cellValue = cell.toString().trim();
                    break;
            }
        }
        return cellValue.trim();
    }


	/**
	 * 获得整个工作表的数据，存入listData和mapData
	 * @return boolean
	 */
	private boolean getSheetData( Sheet sheet) {
		
        listData = new ArrayList<List<String>>();
        mapData = new ArrayList<Map<String, String>>();    
        columnHeaderList = new ArrayList<String>();
        numOfColumn = 0 ;
        
        try {
	        numOfRows = sheet.getLastRowNum() + 1;        
	        for (int i = 0; i < numOfRows; i++) {
	            Row row = sheet.getRow(i);
	            Map<String, String> map = new HashMap<String, String>();
	            List<String> list = new ArrayList<String>();
	            if (row != null) {
	            	if( numOfColumn < row.getLastCellNum() ) numOfColumn = row.getLastCellNum() ;
	                for (int j = 0; j < row.getLastCellNum(); j++) {
	                    Cell cell = row.getCell(j);
	                    if (i == 0){
	                        columnHeaderList.add(getCellValue(cell));
	                    }
	                    else{                        
	                        map.put(columnHeaderList.get(j), this.getCellValue(cell));
	                    }
	                    list.add(this.getCellValue(cell));
	                }
	            }
	            if (i > 0){
	                mapData.add(map);
	            }
	            listData.add(list);
	        }
	        
	        flag = true;
    	
        }catch( Exception e) {
        	e.printStackTrace();
        }
        
        return flag;
    }


	/**
	 * 打印第一个工作表的表头名
	 */
	public void printColumnHeades() {
		this.printColumnHeades(0);
	}

	/**
	 * 按工作表序列打印表头名
	 */
	public void printColumnHeades(int index) {
		
		if( index >= workBook.getNumberOfSheets()) {
        	System.err.println("Error: index greater than NumberOfSheets !");
        	return ;
		}
		
		Sheet sheet =  workBook.getSheetAt(index);
		
		if( sheet.getLastRowNum() < 1 ) {
			System.out.println("INFO: Table index:"+index+" have no content !");
			return ;
		}
			
		Row row =sheet.getRow(0);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			Cell cell = row.getCell(i);
			System.out.println(i + " -- " + getCellValue(cell) );	
		}
	}
	
	/**
	 * 获得表头名列表
	 * @return List<String>
	 */
	public List<String> getColumnHeaderList() {
		return columnHeaderList;
	}

	/**
	 * 获取工作表数量
	 * @return int
	 */
	public int getNumberOfSheets() {
    	return workBook.getNumberOfSheets() ;
    }
    
	/**
	 * 以工作表序列读取工作表内容
	 * @return boolean
	 */
	public boolean getSheetDataByIndex(int index) {
		
		if( index >= workBook.getNumberOfSheets()) {
        	System.err.println("Error: index greater than NumberOfSheets !");
        	return false;
		}
		
		return this.getSheetData( workBook.getSheetAt(index) ) ;
    }

	/**
	 * 以工作表名读取工作表内容
	 * @return boolean
	 */
	public boolean getSheetDataByName(String name) {
    	return this.getSheetData(workBook.getSheet(name)) ;
    }

	/**
	 * 获取工作表行数
	 * @return int
	 */
	public int getNuberOfRows() {
        if(!flag){
        	System.err.println("Error: to run getSheetDataByIndex or getSheetDataByName First !");
        }    
		return this.numOfRows ;
	}
	
	/**
	 * 获取工作表列数
	 * @return int
	 */
	public int getNuberOfColumns() {
        if(!flag){
        	System.err.println("Error: to run getSheetDataByIndex or getSheetDataByName First !");
        }    
		return this.numOfColumn ;
	}

	/**
	 * 以列号和选号获取单元格内容
	 * @param row [ 0 ~ numOfRows-1 ]
	 * @param col [ 1 ~ numOfColumn ]
	 * @return String
	 */
    public String getCellData(int row, int col){
        if(row<0 || col<=0 || row>this.numOfRows || col>this.numOfColumn){
            return null;
        }
        if(!flag){
        	System.err.println("Error: to run getSheetDataByIndex or getSheetDataByName First !");
        	return null;
        }        
        if(listData.size()>=row && listData.get(row).size()>=col){
            return listData.get(row).get(col-1);
        }else{
            return null;
        }
    }

	/**
	 * 以列号和表头名称获取单元格内容
	 * @param row   range: 1 ~ numOfRows-1
	 * @return String
	 */
    public String getCellData(int row, String headerName){
        if(row<=0 || row>this.numOfRows-1 ){
            return null;
        }
        if(!flag){
        	System.err.println("Error: to run getSheetDataByIndex or getSheetDataByName First !");
        	return null;
        }        
        if(mapData.size()>=row && mapData.get(row-1).containsKey(headerName)){
            return mapData.get(row-1).get(headerName);
        }else{
            return null;
        }
    }

	/**
	 * 打印工作表所有内容
	 */
    public void printSheetAllContent() {
    	for( int i = 0 ; i < this.numOfRows ; i++) {
    		for( int j = 1 ; j<= this.numOfColumn ; j++) {
    			System.out.print( getCellData(i , j )+ "\t") ;
    		}
    		System.out.println("");
    	}
    }

    
//    public static void main(String[] args) {
//        ExcelReader eh = new ExcelReader("d:\\test.xls");
//        if( eh.getSheetDataByIndex(0) ) {     	
//	        System.out.println(eh.getNumberOfSheets() +" === " +eh.getNuberOfRows() + " === " +eh.getNuberOfColumns() ) ;
//	        System.out.println(eh.getCellData(0,1));
////	        System.out.println(eh.getCellData(1, "strB1"));
//        	eh.printSheetAllContent();
//	        System.out.println(eh.getCellData(1,1));
//	        System.out.println(eh.getCellData(1,"姓名"));
//	        
//        }
//    }
    
}