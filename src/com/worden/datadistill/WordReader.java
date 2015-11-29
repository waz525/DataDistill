package com.worden.datadistill;

import java.io.*;
import java.util.*;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class WordReader {
	private Range range ;
    private List<String> columnHeaderList = null;
    private List<List<String>> listData;
    private List<Map<String,String>> mapData;
    private boolean flag = false;
    private int numOfRows = 0 ;
    private int numOfColumn = 0 ;
    private int numOfTables = 0 ;
	
	
	public WordReader(String filePath) {
		this.LoadWordFile(filePath);
	}

	/**
	 * 加载Word文件
	 */
	public void LoadWordFile(String filePath) {
        FileInputStream in;
		try {
			in = new FileInputStream(filePath);
			POIFSFileSystem pfs = new POIFSFileSystem(in);     
			HWPFDocument hwpf = new HWPFDocument(pfs);
        	range = hwpf.getRange();//得到文档的读取范围 
		} catch (Exception e) {
			e.printStackTrace();
		}//载入文档     
	}
	

	/**
	 * 获取工作表数量
	 * @return int
	 */
	public int getNumberOfTables() {

        return numOfTables ;
    }
	
	/**
	 * 获取表内容，此为操作函数
	 * @return boolean
	 */	
	public boolean getTableContent() {
		getTableNumbers() ;
		return getTableContent(0) ;
	}
	
	/**
	 * 获取工作表数量
	 * @return int
	 */
	private int getTableNumbers() {
		numOfTables = 0 ;
        try{

        	TableIterator it = new TableIterator(range); 
        	while( it.hasNext() ) {
        		numOfTables++ ;
        		it.next() ;
        		
        	}
        }catch( Exception e) {
        	e.printStackTrace();
        }
        return numOfTables ;
	}
    
	/**
	 * 获取表内容
	 * @param index 表序列值
	 * @return boolean
	 */
	private boolean getTableContent(int index) {
		
        listData = new ArrayList<List<String>>();
        mapData = new ArrayList<Map<String, String>>();    
        columnHeaderList = new ArrayList<String>();
        numOfColumn = 0 ;
        this.numOfRows = 0 ;
        
        if( index < 0 || index > this.numOfTables) {
        	System.err.println("getTableContent: index is out of range !");
        	return false ;
        }

        try{
        	TableIterator it = new TableIterator(range); 
            Table tb = it.next() ;
            int ind = 0 ;
           //迭代文档中的表格  
            while (ind < index) {
                tb = it.next(); 
                ind++ ;
            }
            numOfRows =  tb.numRows() ;
            //迭代行，默认从0开始  
            for (int i = 0; i < tb.numRows(); i++) {  
            	TableRow tr = tb.getRow(i);
                Map<String, String> map = new HashMap<String, String>();
                List<String> list = new ArrayList<String>(); 
	            
	            if ( numOfColumn < tr.numCells() ) numOfColumn = tr.numCells() ;
	            //迭代列，默认从0开始 
	            for (int j = 0; j < tr.numCells(); j++) {  
	            	TableCell td = tr.getCell(j);//取得单元格  
	            	//取得单元格的内容 
	            	Paragraph para =td.getParagraph(0);  
	            	String cellData = para.text();
	            	cellData = cellData.substring(0, cellData.length()-1) ;

	            	
	           	 	if (i == 0){
	           	 		columnHeaderList.add(cellData);
	           	 	}
	           	 	else{
	           	 		map.put(columnHeaderList.get(j), cellData);
	           	 	}
	           	 	list.add(cellData);
	           	  
	            }//end for 
	
	           	if (i > 0){
	           		mapData.add(map);
	           	}
	           	listData.add(list);
	           	
            }//end for
            flag = true ;
        }catch(Exception e){  
            e.printStackTrace();  
        }
        
		return flag ;
	}

	/**
	 * 以工作表序列读取工作表内容
	 * @return boolean
	 */
	public boolean getTableDataByIndex(int index) {
		
		return this.getTableContent(index);
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
	 * @return String
	 */
    public String getCellData(int row, int col){
        if(row<0 || col<0 || row>=this.numOfRows || col>this.numOfColumn){
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
    public void printTableContent() {
    	for( int i = 0 ; i < this.numOfRows ; i++) {
    		for( int j = 1 ; j<= this.numOfColumn ; j++) {
    			System.out.print( getCellData(i , j )+ "\t") ;
    		}
    		System.out.println("");
    	}
    }

//    public static void main(String[] args) {
//    	WordReader eh = new WordReader("d:\\test.doc");
//    	eh.getTableContent() ;
//    	System.out.println("Table Number : "+eh.getNumberOfTables());
//    	System.out.println(eh.getNuberOfRows() + " === " +eh.getNuberOfColumns() ) ;
//    	eh.printTableContent();
//    	System.out.println(eh.getCellData(3, "手机"));
//    	System.out.println(eh.getCellData(2, 1));
//        
//    }
}
