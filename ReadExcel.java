package com.tuma.util;

import java.io.BufferedInputStream;
import java.io.BufferedReader;
import java.io.File;

import java.io.FileInputStream;

import java.io.FileNotFoundException;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.DecimalFormat;

import java.text.SimpleDateFormat;

import java.util.ArrayList;

import java.util.Arrays;

import java.util.Date;

import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;

import org.apache.poi.hssf.usermodel.HSSFRow;

import org.apache.poi.hssf.usermodel.HSSFSheet;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ReadExcel {

	public static void main(String[] args) throws Exception {
//		File file = new File("E:\\下载.xls");
//		readExcel(file);	
			
//		String[][] data = getData(new File("C:\\Users\\Admin\\Desktop\\demo.xls"), 0);
//		for (int i = 0; i < data.length; i++) {
//			System.out.println(Arrays.toString(data[i]));
//		}
		
		
//		String str = openFile("C:\\Users\\Admin\\Desktop\\项目内容进程记录.txt");
//		System.out.println(str);
//		Pattern pattern = Pattern.compile("<td( .*?)?>.*?</td>");
//		Matcher matcher = pattern.matcher(str);
//		List<String> list = new ArrayList<String>();
//		while (matcher.find()) {
//			list.add(matcher.group());
//			System.out.println(matcher.group());
//		}
		
//		File file = new File("C:\\Users\\Admin\\Desktop\\demo.xlsx");
//		
//		if(file.isFile()) {
//			System.out.println(file);
//		}
//
//		String[][] result = getData(file, 1);
//
//		int rowLength = result.length;
//
//		for (int i = 0; i < rowLength; i++) {
//			for (int j = 0; j < result[i].length; j++) {
//				System.out.print(result[i][j].trim() + "\t\t");
//			}
//			System.out.println();
//		}
	 
//		Date date=new Date();
//		long ld=date.getTime()/(24*3600);
//		System.out.println(ld);
//		Date d=new Date(41763*24*3600);
//		SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd HH:mm");
//		System.out.println(sdf.format(d));
//		System.out.println(sdf.format(new Date(ld)));
	}
	private static String ENCODE = "UTF-8";
    public static String openFile( String szFileName ) {
        try {
            BufferedReader bis = new BufferedReader(new InputStreamReader(new FileInputStream( new File(szFileName)), ENCODE) );
            String szContent="";
            String szTemp;
            while ( (szTemp = bis.readLine()) != null) {
                szContent+=szTemp+"\n";
            }
            bis.close();
            return szContent;
        }
        catch( Exception e ) {
            return "";
        }
    }
	// 去读Excel的方法readExcel，该方法的入口参数为一个File对象  
    public static void readExcel(File file) {  
        try {  
            // 创建输入流，读取Excel  
            InputStream is = new FileInputStream(file.getAbsolutePath());  
            // jxl提供的Workbook类  
            Workbook wb = Workbook.getWorkbook(is);  
            // Excel的页签数量  
            int sheet_size = wb.getNumberOfSheets();  
            System.out.println(sheet_size);
            for (int index = 0; index < sheet_size; index++) {
                // 每个页签创建一个Sheet对象  
                Sheet sheet = wb.getSheet(index);  
                // sheet.getRows()返回该页的总行数  
                for (int i = 0; i < sheet.getRows(); i++) {  
                    // sheet.getColumns()返回该页的总列数  
                    for (int j = 0; j < sheet.getColumns(); j++) {  
                        String cellinfo = sheet.getCell(j, i).getContents();  
                        System.out.print(cellinfo + "\t");  
                    }  
                    System.out.println();
                }  
            }  
        } catch (FileNotFoundException e) {  
            e.printStackTrace();  
        } catch (BiffException e) {  
            e.printStackTrace();  
        } catch (IOException e) {  
            e.printStackTrace();  
        }  
    }  
	/**
	 * 
	 * 读取Excel的内容，第一维数组存储的是一行中格列的值，二维数组存储的是多少个行
	 * 
	 * @param file
	 *            读取数据的源Excel
	 * 
	 * @param ignoreRows
	 *            读取数据忽略的行数，比喻行头不需要读入 忽略的行数为1
	 * 
	 * @return 读出的Excel中数据的内容
	 * 
	 * @throws FileNotFoundException
	 * 
	 * @throws IOException
	 */

	public static String[][] getData(File file, int ignoreRows) throws FileNotFoundException, IOException {

		List<String[]> result = new ArrayList<String[]>();

		int rowSize = 0;

		BufferedInputStream in = new BufferedInputStream(new FileInputStream(file));

		// 打开HSSFWorkbook

		POIFSFileSystem fs = new POIFSFileSystem(in);

		HSSFWorkbook wb = new HSSFWorkbook(fs);

		HSSFCell cell = null;

		for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {

			HSSFSheet st = wb.getSheetAt(sheetIndex);

			// 第一行为标题，不取

			for (int rowIndex = ignoreRows; rowIndex <= st.getLastRowNum(); rowIndex++) {

				HSSFRow row = st.getRow(rowIndex);

				if (row == null) {

					continue;

				}

				int tempRowSize = row.getLastCellNum() + 1;
				
				if (tempRowSize > rowSize) {

					rowSize = tempRowSize;

				}

				String[] values = new String[rowSize];

				Arrays.fill(values, "");

				boolean hasValue = false;

				for (short columnIndex = 0; columnIndex <= row.getLastCellNum(); columnIndex++) {

					String value = "";

					cell = row.getCell(columnIndex);

					if (cell != null) {

						// 注意：一定要设成这个，否则可能会出现乱码

						cell.setEncoding(HSSFCell.ENCODING_UTF_16);

						switch (cell.getCellType()) {

						case HSSFCell.CELL_TYPE_STRING:

							value = cell.getStringCellValue();

							break;

						case HSSFCell.CELL_TYPE_NUMERIC:

							if (HSSFDateUtil.isCellDateFormatted(cell)) {

								Date date = cell.getDateCellValue();

								if (date != null) {

									value = new SimpleDateFormat("yyyy-MM-dd").format(date);

								} else {

									value = "";

								}

							} else {
								if (columnIndex == 1) {// 取成交价格
									value = new DecimalFormat("0.00").format(cell.getNumericCellValue());
									
								}else {
									
									value = new DecimalFormat("00.000000").format(cell.getNumericCellValue());
									
									if(Double.valueOf(value)>1000){//--  不是经纬度
										value = new DecimalFormat("0").format(cell.getNumericCellValue());
									}
								}

							}

							break;

						case HSSFCell.CELL_TYPE_FORMULA:

							// 导入时如果为公式生成的数据则无值

							if (!cell.getStringCellValue().equals("")) {

								value = cell.getStringCellValue();

							} else {

								value = cell.getNumericCellValue() + "";

							}

							break;

						case HSSFCell.CELL_TYPE_BLANK:

							break;

						case HSSFCell.CELL_TYPE_ERROR:

							value = "";

							break;

						case HSSFCell.CELL_TYPE_BOOLEAN:

							value = (cell.getBooleanCellValue() == true ? "Y"

							: "N");

							break;

						default:

							value = "";

						}

					}

					if (columnIndex == 0 && value.trim().equals("")) {

						break;

					}

					values[columnIndex] = rightTrim(value);

					hasValue = true;

				}

				if (hasValue) {

					result.add(values);

				}

			}

		}

		in.close();

		String[][] returnArray = new String[result.size()][rowSize];

		for (int i = 0; i < returnArray.length; i++) {

			returnArray[i] = (String[]) result.get(i);

		}

		return returnArray;

	}

	/**
	 * 
	 * 去掉字符串右边的空格
	 * 
	 * @param str
	 *            要处理的字符串
	 * 
	 * @return 处理后的字符串
	 */

	public static String rightTrim(String str) {

		if (str == null) {

			return "";

		}

		int length = str.length();

		for (int i = length - 1; i >= 0; i--) {

			if (str.charAt(i) != 0x20) {

				break;

			}

			length--;

		}

		return str.substring(0, length);

	}
}
//我觉得是那么个意思！！！
