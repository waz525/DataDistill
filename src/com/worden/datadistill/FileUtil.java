package com.worden.datadistill;

import java.io.*;

public class FileUtil
{
        public static void WriteLog( String filename , String content )
        {
                try
                {
                        File outFile = new File( filename ) ;
                        FileWriter fwstr = new FileWriter( outFile  , true ) ;
                        String str = "" ;
                        java.util.Date nowDate = new java.util.Date() ;
                        java.text.SimpleDateFormat formatter = new java.text.SimpleDateFormat("[yyyy-MM-dd HH:mm:ss]") ;
                        String datestr = formatter.format(nowDate);
                        str = datestr + ": " + content +" \n";
                        fwstr.write( str ) ;
                        fwstr.close();
                }
                catch( IOException e )
                {
			e.printStackTrace();
                        System.out.println(e);
                }
        }
	
	public static void WriteString( String filename , String content )
	{
		try
		{
			File outFile = new File( filename ) ;
			FileWriter fwstr = new FileWriter( outFile  , true ) ;
			fwstr.write( content ) ;
			fwstr.close();
		}
		catch( IOException e )
		{
			e.printStackTrace();
			System.out.println(e);
		}
	}

	public static String ReadString( String filename )
	{
		String rst = "" ;
		Reader reader = null;
		try
		{
			int charread = 0;
			reader = new InputStreamReader(new FileInputStream(filename));
			while ((charread = reader.read()) != -1)
			{
				rst = rst + (char)charread ;
			}
			reader.close();
		}
		catch( IOException e )
		{
			e.printStackTrace();
			System.out.println(e);
		}
		finally 
		{
			if (reader != null) 
			{
				try 
				{ 
					reader.close();
				} 
				catch (IOException e1) 
				{
					e1.printStackTrace();
				}
			}
		}
		return rst ;
	}
	
	public static void RemoveFile( String filename ) 
	{
		File file = new File(filename);
		if (file.isFile() && file.exists())
		{
			file.delete();
		}
	}

	public static boolean isFile( String filename )
	{
		File file = new File(filename);
		if( file.isFile() )
		{
			return true ;
		}
		else
		{
			return false ;
		}
	}
}

