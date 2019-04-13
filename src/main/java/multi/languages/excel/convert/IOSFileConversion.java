package multi.languages.excel.convert;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class IOSFileConversion {
	
	
	public static void main(String args[]) {
//        readExcel("d:/IOS.xls");
        readFile("d:/IOS.txt");//读取文件，同时生成Excel模板
    }
	
	public static void readFile(String filePath) {
		
      try (FileReader reader = new FileReader(filePath);
           BufferedReader br = new BufferedReader(reader) // 建立一个对象，它把文件内容转成计算机能读懂的语言
        ) {
          String line;
          //网友推荐更加简洁的写法
          List<Map<String,String>> list = new ArrayList<>();
          while ((line = br.readLine()) != null) {
             if(line.contains("=")) {
          		Map<String,String> map = new HashMap<String, String>();
          		String value[] = line.split("=");
          		String realKey= value[0].trim().replace("\"", "");
          		String realValue= value[1].replace(";", "").replace("\"", "");
          		map.put(realKey, realValue);
          		list.add(map);
          	}
          }
          createExcel(list);//生产EXcel
      } catch (IOException e) {
          e.printStackTrace();
      }
  }
	
	public static void writeFile(List<Map<String,String>>list) {
        try {
            File writeName = new File("d:/IOS.txt"); // 相对路径，如果没有则要建立一个新的output.txt文件
            writeName.createNewFile(); // 创建新文件,有同名的文件的话直接覆盖
            try (FileWriter writer = new FileWriter(writeName);
                 BufferedWriter out = new BufferedWriter(writer)
            ) {
            	for (Map<String,String> map : list) {
                    for (Entry<String, String> entry : map.entrySet()) {
                    	String key = entry.getKey().trim();
                    	String value = entry.getValue().trim();
                    	out.write("\""+key+"\"=\""+value+"\";"+"\r\n");
                    }
                }
                out.flush(); // 把缓存区内容压入文件
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
	
	
    /**
     * 读取Excel
     * @param filePath
     */
    public static void  readExcel(String filePath){
    	
    	Workbook wb =null;
        Sheet sheet = null;
        Row row = null;
        List<Map<String,String>> list = null;
//        String cellData = null;
//        String columns[] = {"key","zh","en","fr"};
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if(".xls".equals(extString)){
                 wb = new HSSFWorkbook(is);
            }else if(".xlsx".equals(extString)){
                 wb = new XSSFWorkbook(is);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        if(wb != null){
            //用来存放表中数据
            list = new ArrayList<Map<String,String>>();
            //获取第一个sheet
            sheet = wb.getSheetAt(0);
            //获取最大行数
            int rownum = sheet.getLastRowNum()+1;
            //获取第一行
            row = sheet.getRow(0);
            //获取最大列数
//            int colnum = row.getPhysicalNumberOfCells();
            for (int i = 1; i<rownum; i++) {
                Map<String,String> map = new LinkedHashMap<String,String>();
                row = sheet.getRow(i);
                if(row !=null){
                	if(null != row.getCell(0)) {
                		System.out.println(i);
                		map.put(row.getCell(0).toString(), getCellFormatValue(row.getCell(1)).toString());
                		list.add(map);
                	}
                }else{
                    continue;
                }
            }
        }
        writeFile(list);
    }
    
    /**
     * 生测Excel
     * @param os
     * @throws WriteException
     * @throws IOException
     */
    public static void createExcel(List<Map<String,String>> list) throws IOException{
    	XSSFWorkbook wb = new XSSFWorkbook();
    	// 建立新的sheet对象（excel的表单）
    	XSSFSheet sheet = wb.createSheet("sheet1");
    	// 在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
    	XSSFRow row0 = sheet.createRow(0);
    	// 添加表头
    	row0.createCell(0).setCellValue("key");
    	row0.createCell(1).setCellValue("value");
    	int line =1;
    	for (Map<String,String> map : list) {
            for (Entry<String, String> entry : map.entrySet()) {
            	XSSFRow row = sheet.createRow(line);
            	row.createCell(0).setCellValue(entry.getKey().trim());
            	row.createCell(1).setCellValue(entry.getValue().trim());
            }
            line++;
        }
        FileOutputStream output=new FileOutputStream("d:/IOS_.xlsx");
        wb.write(output);//写入磁盘  
    	output.close();
    }
    
    public static Object getCellFormatValue(Cell cell){
        Object cellValue = null;
        if(cell!=null){
            //判断cell类型
            switch(cell.getCellType()){
            case Cell.CELL_TYPE_NUMERIC:{
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            }
            case Cell.CELL_TYPE_FORMULA:{
                //判断cell是否为日期格式
                if(DateUtil.isCellDateFormatted(cell)){
                    //转换为日期格式YYYY-mm-dd
                    cellValue = cell.getDateCellValue();
                }else{
                    //数字
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }
                break;
            }
            case Cell.CELL_TYPE_STRING:{
                cellValue = cell.getRichStringCellValue().getString();
                break;
            }
            default:
                cellValue = "";
            }
        }else{
            cellValue = "";
        }
        return cellValue;
    }


}
