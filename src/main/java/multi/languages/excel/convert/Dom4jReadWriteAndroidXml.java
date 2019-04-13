package multi.languages.excel.convert;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

public class Dom4jReadWriteAndroidXml {
	
	public static void main(String[] args) {
		String xmlFilePath="d:/String.xml";
		String outputXmlFilePath="d:/android.xml";
		try {
			List<Map<String,Object>> xmlNodeMapList=readXmlFileContentToMapList(xmlFilePath);
			//System.out.println(xmlNodeMapList);
			writeMapListToXmlFile(xmlNodeMapList,outputXmlFilePath);
		} catch (Exception e) {
			e.printStackTrace();
		}		
	}
		
	public static List<Map<String,Object>> readXmlFileContentToMapList(String xmlFilePath) throws DocumentException {
		File file = new File(xmlFilePath);
        SAXReader reader = new SAXReader();
        List<Map<String,Object>> xmlNodeMapList = new ArrayList<>();
        try {
            Document document = reader.read(file);
            Element root = document.getRootElement();          
           List<Element>  elementList= root.elements();
            /**
            Map<String,Object> mapElement = new HashMap<String, Object>();
            Map<String, Object> resultMap=getAllElements(elementList,mapElement);
            System.out.println(resultMap);
            */
            for (Element everyNode:elementList) {            	
            	String attrName=everyNode.attributeValue("name");
            	String nodeText=everyNode.getText();
            	int elementLen= everyNode.elements().size();
            	Map<String,Object> nodeElementMap = new HashMap<String, Object>();
            	if(elementLen>1) {
            		List<Element>  childNodeList=everyNode.elements();
            		List<String> childNodeValueList=new ArrayList<String>();
            		 for (Element childNode:childNodeList) {
            			 //String childNodeName=childNode.attributeValue("name");
                     	 String childNodeText=childNode.getText();
            			 //System.out.println("childNodeName:"+childNodeName+" childNodeText: "+childNodeText);
                     	childNodeValueList.add(childNodeText);
            		 }  
            		 nodeElementMap.put("attrName", attrName);
                 	 nodeElementMap.put("nodeContent", childNodeValueList);
            	}else {
            		nodeElementMap.put("attrName", attrName);
                	nodeElementMap.put("nodeContent", nodeText);
            	}            	
            	xmlNodeMapList.add(nodeElementMap);
            	//System.out.println("attrName:"+attrName+" nodeText: "+nodeText);
			}
        	return xmlNodeMapList;
        } catch (DocumentException e) {
            e.printStackTrace();
        }
        return null;
	}
	
	/**
	 * 写入mapList信息 到XML文件之中
	 */
	public static void writeMapListToXmlFile(List<Map<String,Object>> mapNodeList,String outputXmlFileNamePath) throws Exception {
		//1.创建一个Document对象
		Document doc = DocumentHelper.createDocument();
 
		//2.创建根对象
		Element rootNodeElement = doc.addElement("resources");
		rootNodeElement.setAttributeValue("xmlns:xliff", "urn:oasis:names:tc:xliff:document:1.2"); 
		
		//3.循环创建子节点 然后添加到xml节点元素之上
		for(Map<String,Object> nodeMapItem: mapNodeList) {
				String attrName=nodeMapItem.get("attrName").toString();
				 Object nodeContent=nodeMapItem.get("nodeContent");
				 Element nodeItemElement =null;
				//判断一下 当前节点是否包含有子节点
				 if( nodeContent instanceof List) {
					 List<String> childrenList =(List<String>)nodeContent;
					if(childrenList.size()>0) {
						 nodeItemElement=rootNodeElement.addElement("string-array");
						 nodeItemElement.setAttributeValue("name", attrName);
						 for(String nodeText:childrenList) {
							 Element childElement =nodeItemElement.addElement("item");
							 childElement.setText(nodeText);
						 }
					}				 
				 }else {
					 nodeItemElement=rootNodeElement.addElement("string");
					 nodeItemElement.setAttributeValue("name", attrName);
					 String nodeText=nodeContent.toString();
					 nodeItemElement.setText(nodeText);
				 }	  
		}		
		//6.设置输出流来生成一个xml文件
		OutputStream os = new FileOutputStream(outputXmlFileNamePath);
		//Format格式输出格式刷
		OutputFormat format = OutputFormat.createPrettyPrint();
		//设置xml编码
		format.setEncoding("utf-8");
 
		//写：传递两个参数一个为输出流表示生成xml文件在哪里
		//另一个参数表示设置xml的格式
		XMLWriter xw = new XMLWriter(os,format);
		//将组合好的xml封装到已经创建好的document对象中，写出真实存在的xml文件中
		xw.write(doc);
		//清空缓存关闭资源
		xw.flush();
		xw.close();
	} 
	
}
