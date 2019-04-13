package multi.languages.excel.convert;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

/**
 * Properties 与 Excel 文件互转
 * @author jianxiapc
 *
 */
public class PropertiesExcelInterconvert {
 
	LinkedHashMap< String, String > propertiesContentToExcelMap = new LinkedHashMap< String, String >(); 
	//创建一个HashMap，它将存储键和值从xls文件提供
	LinkedHashMap<String, String > excelContentToPropertiesMap = new LinkedHashMap< String, String >();
	
    public static void main(String[] args) { 
    	PropertiesExcelInterconvert propertiesExcelInterconvert=new PropertiesExcelInterconvert();
    	try {
        	Resource resource = new ClassPathResource("/i18n/messages_zh_CN.properties");
			File propertieFile =  resource.getFile();
			System.out.println(propertieFile.getPath());			
			System.out.println(propertieFile.getName());
			
			String fileDirPath=propertieFile.getPath().replace(propertieFile.getName(), "");
			//System.out.println("fileDirPath: "+fileDirPath);
			String inputPropertieFilePath=propertieFile.getPath();
			String outExcelFileNamePath=fileDirPath+"/messages_zh.xls";
			System.out.println("从  properties 写入 Excel 文件开始!");
			//propertiesExcelInterconvert.propertiesFileConvertToExcel(inputPropertieFilePath, outExcelFileNamePath);
			System.out.println("从  properties 写入 Excel 文件完成!");
			
			System.out.println("从 Excel 写入properties 文件开始!");
			propertiesExcelInterconvert.excelFileConvertToProperties(outExcelFileNamePath, fileDirPath+"messages_zh_xls.properties");
			System.out.println(" Excel 写入properties 文件 完成!");
		} catch (IOException e) {
			e.printStackTrace();
		} 
    } 
    
   /**
    * properties 转换内容转换为Excel
    */
   private void propertiesFileConvertToExcel(String inputPropertieFilePath,String outExcelFileNamePath) {    	
    	System.setProperty("file.encoding", "UTF-8");     
    	readPropertiesContent(inputPropertieFilePath);
		writePropertiesContentToExcelFile(outExcelFileNamePath); 	
    }
   
   /**
    * Excel文件转换为Properties
    */
   private void excelFileConvertToProperties(String inputExcelFileNamePath,String outPropertiesFileNamePath) {
	   PropertiesExcelInterconvert propertiesExcelInterconvert = new PropertiesExcelInterconvert(); 
       // 通过将xls的位置传递给readExcelFileContent()方法，该方法将把键和值从xls加载到HashMap
	   propertiesExcelInterconvert.readExcelFileContent(inputExcelFileNamePath);

       //通过传递属性文件的位置来调用writeExcelToPropertiesFile方法。这个方法将把hashMap中的键和值存储到属性文件中
	   propertiesExcelInterconvert.writeExcelToPropertiesFile(outPropertiesFileNamePath);
   }
    
    /**
          * 读取 properties文件内容
     * @param propertiesFilePath
     */
    private void readPropertiesContent(String propertiesFilePath) { 
    	
        // 创建包含属性路径的文件对象
        File propertiesFile = new File(propertiesFilePath); 
        // 如果属性文件是一个文件，做下面的事情
        if(propertiesFile.isFile()){
            try{
                // 创建一个FileInputStream来加载属性文件
                FileInputStream fisProp = new FileInputStream(propertiesFile);  
                BufferedReader in = new BufferedReader(new InputStreamReader(fisProp, "UTF8"));
                
                // 创建Properties对象并加载 通过FileInputStream将属性键和值赋给它
                // 注意事项：默认的Properties 
                /**
                Java 的 Properties 加载属性文件后是无法保证输出的顺序与文件中一致的，
                                 因为 Properties 是继承自 Hashtable 的， key/value 都是直接存在 Hashtable 中的，
                                 而 Hashtable 是不保证进出顺序的。 此处覆盖原来Properties 写新的 OrderedProperties
                 */
                Properties properties = new OrderedProperties();
                properties.load(in);                
                
                Enumeration< Object > keysEnum = properties.keys();  
                properties.keySet().iterator();                  
            
                while(keysEnum.hasMoreElements()){
                    String propKey = (String)keysEnum.nextElement();
                    String propValue = properties.getProperty(propKey); 
                   
                    Map<String,String> propItem=new HashMap<String,String>();
                    propItem.put(propKey.trim(), propValue.trim());
                    propertiesContentToExcelMap.put( propKey.trim(),propValue.trim());
 
                }    
                // 属性键和值，通过fileinputstreamprint HashMap并关闭文件FileInputStream
                System.out.println("Properties Map ... \n" +  propertiesContentToExcelMap);
                fisProp.close();
 
            }catch(FileNotFoundException e){                     
                e.printStackTrace();
            }
            catch(IOException e){                  
                e.printStackTrace();
            } 
        } 
    }
    
    /**
     * Properties内容写入Excel之中
     * @param excelPath
     */
    private void writePropertiesContentToExcelFile(String excelPath) {
 
        HSSFWorkbook workBook = new HSSFWorkbook();
        
        //创建一个名为Properties 的sheet
        HSSFSheet worksheet = workBook.createSheet("Properties");

        // 在当前sheet中创建第一行
        HSSFRow row = worksheet.createRow((short) 0);
        
        //设置列头样式
        HSSFCellStyle cellStyle = workBook.createCellStyle();      
        cellStyle.setFillForegroundColor(HSSFColor.GOLD.index);
        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
 
        HSSFCell cell1 = row.createCell(0); 
        //设置第一行第一列名称
        cell1.setCellValue(new HSSFRichTextString("Keys")); 
        cell1.setCellStyle(cellStyle);
 
        HSSFCell cell2 = row.createCell(1);
        cell2.setCellValue(new HSSFRichTextString("Values"));                
        cell2.setCellStyle(cellStyle); 

        //循环把 Properties文件内容一行一行添加到Excel之中
        for (String s : propertiesContentToExcelMap.keySet()) {
        	
            //在sheet之中每次增加一行
            HSSFRow rowOne = worksheet.createRow(worksheet.getLastRowNum() + 1);
            // 在此行之中创建两列
            HSSFCell cellZero = rowOne.createCell(0);
            HSSFCell cellOne = rowOne.createCell(1);

            //从map和set之中提取 key和value值
            String key;
            key = s;
            String value = propertiesContentToExcelMap.get(key);
            // 把提取的值设置到 Excel之中的列
            cellZero.setCellValue(new HSSFRichTextString(key));
            cellOne.setCellValue(new HSSFRichTextString(value));
        }         
        try{ 
            FileOutputStream fosExcel; 
            File fileExcel = new File(excelPath);       
            fosExcel = new FileOutputStream(fileExcel); 
            workBook.write(fosExcel); 
            fosExcel.flush();
            fosExcel.close(); 
        }catch(Exception e){ 
            e.printStackTrace(); 
        }
    }
    
    /**
          * 读取Excel文件
     * @param fileName
     */
    public void readExcelFileContent(String fileName)     {
    	 
        HSSFCell cell1 =null;
        HSSFCell cell2 =null; 
        try{        // 通过传递excel的位置创建FileInputStream
            FileInputStream input = new FileInputStream(new File(fileName));    
            //使用HSSFWorkbook对象创建工作簿
            HSSFWorkbook workBook = new HSSFWorkbook(input); 
            // 通过调用获取位置0处的 sheet   
            HSSFSheet sheet = workBook.getSheet("Properties");    
            // 创建 sheet的行迭代器
             Iterator<Row> rowIterator = sheet.rowIterator(); 
            while(rowIterator.hasNext()){   
                // 通过调用创建对row的引用
                HSSFRow row = (HSSFRow) rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator(); 
                // Iterating over each cell
                while(cellIterator.hasNext()){    
                    cell1 = (HSSFCell) cellIterator.next(); 
                    String key = cell1.getRichStringCellValue().toString(); 
                    if(!cellIterator.hasNext()){ 
                        String value = ""; 
                        //把key和value放置到 properties Map对象之中
                        excelContentToPropertiesMap.put(key, value);  
                    }  else {
                        cell2 = (HSSFCell) cellIterator.next();
                        String value = cell2.getRichStringCellValue().toString();
                        excelContentToPropertiesMap.put(key, value);       
                    }                          
                }         
            }    
        } 
        catch (Exception e){ 
            System.out.println("没有发生此类元素异常 ..... ");
            e.printStackTrace(); 
        }               
    }
    
    /**
     * Excel写回到Properties文件之中
     * @param propertiesPath
     */
    public void writeExcelToPropertiesFile(String propertiesPath) {
    	
        Properties props = new OrderedProperties();
 
        //创建一个文件对象，该对象将指向属性文件的位置
        File propertiesFile = new File(propertiesPath);
 
        try {
 
            //通过传递上述属性文件创建FileOutputStream 并且设置每次覆盖重写
            FileOutputStream xlsFos = new FileOutputStream(propertiesFile,false);
 
            // 首先将哈希映射键转换为Set，然后对其进行迭代。
            Iterator<String> mapIterator = excelContentToPropertiesMap.keySet().iterator();
 
            //遍历迭代器属性
            while(mapIterator.hasNext()) {
 
                // extracting keys and values based on the keys
                String key = mapIterator.next().toString();
 
                String value = excelContentToPropertiesMap.get(key);
                if(!key.equals("Keys")) {
                	//在上面创建的props对象中设置每个属性key与value
                    props.setProperty(key, value);
                }    
            }
 
            //最后将属性存储到实属性文件中。
            props.store(new OutputStreamWriter(xlsFos, "utf-8"), null);
 
        } catch (FileNotFoundException e) {
 
            e.printStackTrace();
 
        } catch (IOException e) {
 
            e.printStackTrace();
 
        }
    }
}
