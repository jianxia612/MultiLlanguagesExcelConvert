package multi.languages.excel.convert;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class AndroidXmlExcelInterconvert {
	
	
	public static void main(String args[]) {
		String xmlFileNamePath="d:/String.xml";
		String outputExcelFileNamePath="d:/android.xlsx";
		String outputXmlFileNamePath="d:/android.xml";
		try {
			//读取xml内容到Excel文件之中
			System.out.println("从  AndroidXml 写入 Excel 文件开始!");
			readAndroidXmlToExcel(xmlFileNamePath,outputExcelFileNamePath);
			System.out.println("从  AndroidXml 写入 Excel 文件完成!");
			//写入Excel文件内容到xml文件之中
			System.out.println("从 Excel 写入AndroidXml 文件开始!");
			readExcelOutputXmlFile(outputExcelFileNamePath, outputXmlFileNamePath);
			System.out.println(" Excel 写入AndroidXml 文件 完成!");
		} catch (Exception e) {
			e.printStackTrace();
		}
    }
	
	public static void writeExcelListToXMLFile(List<Map<String,Object>> excelXmlNodeList,String outputXmlFileNamePath) {
        try {
        	Dom4jReadWriteAndroidXml.writeMapListToXmlFile(excelXmlNodeList, outputXmlFileNamePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
	
	public static void readAndroidXmlToExcel(String xmlFileNamePath,String outputExcelFileNamePath) throws Exception {
		List<Map<String,Object>> xmlNodeMapList=Dom4jReadWriteAndroidXml.readXmlFileContentToMapList(xmlFileNamePath);
		
		List<Map<String,String>> xmlNodeExcelMapList=new ArrayList<Map<String,String>>();
		for(Map<String,Object> xmlNodeMap:xmlNodeMapList) {		
			 String attrName=xmlNodeMap.get("attrName").toString();
			 Object nodeContent=xmlNodeMap.get("nodeContent");
			//判断一下 当前节点是否包含有子节点
			 if( nodeContent instanceof List) {
				continue;			 
			 }else {
				 String nodeText=nodeContent.toString();	
				 Map<String,String> excelItemMap=new HashMap<String,String>();
				 excelItemMap.put("key", attrName);
				 excelItemMap.put("value", nodeText);
				 xmlNodeExcelMapList.add(excelItemMap);
			 }
		}
		readAndroidXmlContentWriteToExcel(xmlNodeExcelMapList,outputExcelFileNamePath);//生产EXcel
  }
	
    /**
     * 读取Excel
     * @param filePath
     */
    public static void  readExcelOutputXmlFile(String excelFilePath,String outputXmlFileNamePath){
    	
    	Workbook wb =null;
        Sheet sheet = null;
        Row row = null;
        List<Map<String,Object>> list = null;
		//String cellData = null;
		//String columns[] = {"key","zh","en","fr"};
        String extString = excelFilePath.substring(excelFilePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(excelFilePath);
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
            list = new ArrayList<Map<String,Object>>();
            //获取第一个sheet
            sheet = wb.getSheetAt(0);
            //获取最大行数
            int rownum = sheet.getLastRowNum()+1;
            //获取第一行
            row = sheet.getRow(0);
            //获取最大列数
//            int colnum = row.getPhysicalNumberOfCells();
            for (int i = 1; i<rownum; i++) {
                Map<String,Object> map = new LinkedHashMap<String,Object>();
                row = sheet.getRow(i);
                if(row !=null){
                	if(null != row.getCell(0)) {
                		String attrName=row.getCell(0).toString();
       				 	Object nodeContent=getCellFormatValue(row.getCell(1)).toString();
                		map.put("attrName",attrName);
                		map.put("nodeContent",nodeContent);                		
                		list.add(map);
                	}
                }else{
                    continue;
                }
            }
        }
        writeExcelListToXMLFile(list,outputXmlFileNamePath);
    }
    
    /**
     * 生测Excel
     * @param os
     * @throws WriteException
     * @throws IOException
     */
    public static void readAndroidXmlContentWriteToExcel(List<Map<String,String>> list,String outputExcelFileNamePath) throws IOException{
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
    		XSSFRow row = sheet.createRow(line);
        	row.createCell(0).setCellValue(map.get("key").trim());
        	row.createCell(1).setCellValue(map.get("value").trim());
            line++;
        }
        FileOutputStream output=new FileOutputStream(outputExcelFileNamePath);
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
