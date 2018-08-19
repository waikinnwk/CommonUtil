package com.fds.focal.docx;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.apache.commons.codec.binary.Base64;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.finders.RangeFinder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTMarkupRange;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;

import com.fds.util.Util;


public class DOCXMailMergeUtil {
	final static String XPATH_TO_SELECT_TEXT_NODES = "//w:t";
			
	public static void main(String[] args) throws Exception {
		
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(
				new File(
						"Y:/99-Workarea/9903-Share/DocumentTemplate/NoticeofPayment_Dental_CSSA.docx"));
		
		Map<String,Object> stringMap = new HashMap();

		
		stringMap.put("clinicEmailForPay", "dsdsd");
		
		stringMap.put("clinicNameCCF","dsdsdsdsd");
		stringMap.put("clinicContactTitle", "AAAAAA");
		stringMap.put("clinicContactName","AAAAAAAAAA");
		
		stringMap.put("clinicNameCCF2","AAAAAAAAAA");
		stringMap.put("clinicBankAccNo","AAAAAAAAAA");
		stringMap.put("clinicBankAccName","AAAAAAAAAA");
		List list = new ArrayList<Map>();
		Map subMap = new HashMap();
		subMap.put("appNo", "1");
		list.add(subMap);
		
		subMap = new HashMap();
		subMap.put("appNo", "2");
		list.add(subMap);
		
		subMap = new HashMap();
		subMap.put("appNo", "3");
		list.add(subMap);
		
		subMap = new HashMap();
		subMap.put("appNo", "4");
		list.add(subMap);
		
		stringMap.put("appList",list);
		
				/*
		map = new HashMap<DataFieldName, String>();
		map.put( new DataFieldName("Kundenname"), "Jason");
		map.put(new DataFieldName("Kundenstrasse"), "Collins Street");
		
		data.add(map);		
		*/
		
//		System.out.println(XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getJaxbElement(), true, true));
		//replacePlaceholders(wordMLPackage,stringMap);
		replaceParagraph(stringMap,wordMLPackage);
		//WordprocessingMLPackage output = org.docx4j.model.fields.merge.MailMerger.getConsolidatedResultCrude(wordMLPackage, data);

		//System.out.println(XmlUtils.marshaltoString(output.getMainDocumentPart().getJaxbElement(), true, true));
		
		wordMLPackage.save(new java.io.File(
				"F:/test_11_out.docx") );
		
		
		
		
	}
	
	public static void mailMerge(Map<String,Object> dataMap,String templatePath,ServletOutputStream out) throws Exception{
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(
				new File(templatePath));
		replaceParagraph(dataMap,wordMLPackage);
		wordMLPackage.save(out);
	}
	
	public static String mailMerge(Map<String,Object> dataMap,String templatePath) throws Exception{
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		mailMerge(dataMap, templatePath, out);
		out.close(); 			
		Base64 encoder = new Base64();
		return new String(encoder.encode(out.toByteArray()));
	}
	
	public static void mailMerge(Map<String,Object> dataMap,String templatePath,ByteArrayOutputStream out) throws Exception{
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(
				new File(templatePath));
		
		replaceParagraph(dataMap,wordMLPackage);
		wordMLPackage.save(out);
	}
	
	private static void replacePlaceholders(WordprocessingMLPackage targetDocument,Map templateProperties) throws JAXBException, Exception {

	    List texts = targetDocument.getMainDocumentPart()
	         .getJAXBNodesViaXPath(XPATH_TO_SELECT_TEXT_NODES, true);

	    for (Object obj : texts) {
	         Text text = (Text) ((JAXBElement) obj).getValue();

	         String textValue = text.getValue();
	         for (Object key : templateProperties.keySet()) {
	              textValue = textValue.replaceAll("\\$\\{" + key + "\\}",
	              (String) templateProperties.get(key));
	         }

	         text.setValue(textValue);
	    }
	}
	
	private static void replaceParagraph(Map<String, Object> map, WordprocessingMLPackage template) {

		org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
        List<Object> paragraphs = getAllElementFromObject(template.getMainDocumentPart(), P.class);
              
        
        ArrayList dynamicListParamName = new ArrayList();
        Iterator it = map.entrySet().iterator();
        while (it.hasNext()) {
        	Map.Entry pair = (Map.Entry)it.next();
        	String paramKey = (String) pair.getKey();
        	if(paramKey.endsWith("List")){
        		dynamicListParamName.add(paramKey);
        	}
        } 
        
        for(int i = 0 ; i < dynamicListParamName.size() ; i++){
        	
        	List<Map> list = (List) map.get(dynamicListParamName.get(i));
        	
	        List<Object> tables = getAllElementFromObject(template.getMainDocumentPart(), Tbl.class);
	        for (Object tbl : tables) {  
	        	Tbl tableObj = (Tbl) tbl;
	        	Tr clone = null;
	        	boolean targetTable = false;
	        	boolean withFooter = false;
	        	int rowCount = 0;
	        	int footerRowDiff = 1;
	        	for(Object tr :tableObj.getContent()){
	        		if(tr instanceof Tr){
		        		Tr TrObj = (Tr) tr;
		        		rowCount += 1;
		        		int trCheckCompleted = 0b11;
		        		boolean trAllBookmarkFound = true;
		        		boolean trAtLeastOneBookmarkFound = false;
		        		if(list.size() > 0){
		        			Iterator it2 = list.get(0).entrySet().iterator();
		        		    while(it2.hasNext()) {
		        		        Map.Entry pair2 = (Map.Entry)it2.next();
		        		        String subBookmarkName = (String) pair2.getKey();
		        		        if(!bookMarkFoundInTableRow(TrObj,subBookmarkName)){
		        		        	trAllBookmarkFound = false;
		        		        	trCheckCompleted = trCheckCompleted & 0b10;
		        		        	if (trCheckCompleted==0)	break;
		        		        } else {
		        		        	trAtLeastOneBookmarkFound = true;
		        		        	trCheckCompleted = trCheckCompleted & 0b01;
		        		        	if (trCheckCompleted==0)	break;
		        		        }
		        		    }
		        		}
		
		        		// TODO: Replace Bookmarks in the Tables from HashMap
		        		// Since some bookmark is removed, system should allows other bookmark to be replaced.
		        		if (trAtLeastOneBookmarkFound &&
		        				list.size() > 0){
		        			targetTable = true;
		        			if(rowCount < tableObj.getContent().size()){
		        				withFooter = true;
		        			}
		        			if(Util.isEmpty(clone))
		        				clone = XmlUtils.deepCopy(TrObj); 			
		        			Iterator it2 = list.get(0).entrySet().iterator();
		        		    while(it2.hasNext()) {
		        		        Map.Entry pair2 = (Map.Entry)it2.next();
		        		        String subBookmarkName = (String) pair2.getKey();
		        		        
		    	        		for(Object tc : TrObj.getContent()){
		    	        			if(tc instanceof JAXBElement){
		    	        				JAXBElement jax1 = (JAXBElement) tc;
		    	        				Tc tcObj = (Tc) jax1.getValue();
		    	        				List<Object> tcList = ((ContentAccessor)(tcObj)).getContent();
		    	        				for(int j = 0 ; j < tcList.size(); j++){
		    	        					if (tcList.get(j) instanceof P) {
		    	        						List theList = ((ContentAccessor)(tcList.get(j))).getContent(); 
		    	        						replaceBookmarkInP(factory, theList, subBookmarkName ,list.get(0).get(subBookmarkName),false);
		    	        					}
		    	        				}
		    	        			}
		    	        		}	        		        
	
		        		    }
		        			
		        		    
		        		}
	        		
	        		}
	        		else{
	        			footerRowDiff++;
	        		}
	        	}
	        	if(targetTable && clone != null){
	        		for(int k = 1 ; k < list.size() ; k++){
		        		Tr newTr = XmlUtils.deepCopy(clone);
		        		
		        		
	        			Iterator it2 = list.get(k).entrySet().iterator();
	        		    while(it2.hasNext()) {
	        		        Map.Entry pair2 = (Map.Entry)it2.next();
	        		        String subBookmarkName = (String) pair2.getKey();
	        		        
	    	        		for(Object tc : newTr.getContent()){
	    	        			if(tc instanceof JAXBElement){
	    	        				JAXBElement jax1 = (JAXBElement) tc;
	    	        				Tc tcObj = (Tc) jax1.getValue();
	    	        				List<Object> tcList = ((ContentAccessor)(tcObj)).getContent();
	    	        				for(int j = 0 ; j < tcList.size(); j++){
	    	        					if (tcList.get(j) instanceof P) {
	    	        						List theList = ((ContentAccessor)(tcList.get(j))).getContent(); 
	    	        						replaceBookmarkInP(factory, theList, subBookmarkName ,list.get(k).get(subBookmarkName),false);
	    	        					}
	    	        				}
	    	        			}
	    	        		}	        		        
	
	        		    }
	        		    if(withFooter)
	        		    	tableObj.getContent().add(tableObj.getContent().size() - footerRowDiff, newTr);
	        		    else
	        		    	tableObj.getContent().add(newTr);
	        		}
	        		
	        	}
	        }

        }
        
        for (Object p : paragraphs) {  
            RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
            new TraversalUtil(p, rt);
            
 
        	
            
            
            for (CTBookmark content : rt.getStarts()) {  
                if ((map.get(content.getName()) != null)) {  
                	Object textToAdd = map.get(content.getName());
                	                	
                    List<Object> theList = null;
                    if (content.getParent() instanceof P) {
                    	theList = ((ContentAccessor)(content.getParent())).getContent();
                    } else {
                    	continue; 
                    }
                    if(theList.size() > 0 ){
                    	replaceBookmarkInP(factory, theList,content.getName(), textToAdd,true);
                    }
                    else {
                    	if(textToAdd instanceof String){
	                        R run = factory.createR();
	                        Text t2 = factory.createText();
	                        run.getContent().add(t2);   
	                        t2.setValue((String) textToAdd);
	                        theList.add(0, run); 
                    	}
                    }
                } 
            }  

        } 

    }	
	
	private static boolean bookMarkFound(List<Object> theList, String bookmarkName){
		boolean bookMarkFound = false;
		for(int i = 0 ; i < theList.size() ; i++){
			if(theList.get(i) instanceof JAXBElement ){
    			JAXBElement jax1 = (JAXBElement) theList.get(i);
    			if(jax1.getValue() instanceof CTBookmark){
    				CTBookmark ct = (CTBookmark) jax1.getValue();
    				if(ct.getName().equals(bookmarkName)){
    					bookMarkFound = true;
    					break;
    				}
    			} 	
			}
		}
		return bookMarkFound;
	}
	
	private static boolean bookMarkFoundInTableRow(Tr TrObj,String bookMarkName){
		boolean bookMarkFound = false;
		for(Object tc : TrObj.getContent()){
			if(tc instanceof JAXBElement){
				JAXBElement jax1 = (JAXBElement) tc;
				Tc tcObj = (Tc) jax1.getValue();
				List<Object> tcList = ((ContentAccessor)(tcObj)).getContent();
				for(int j = 0 ; j < tcList.size(); j++){
					if (tcList.get(j) instanceof P) {
						List theList = ((ContentAccessor)(tcList.get(j))).getContent(); 
						if(bookMarkFound(theList,bookMarkName)){
							bookMarkFound = true;
							break;
						}
					}
				}
			}
		}		
		return bookMarkFound;
	}
	
	private static void replaceBookmarkInP(org.docx4j.wml.ObjectFactory factory, List<Object> theList, String bookmarkName, Object textToAdd, boolean forceSet){
       	boolean bookmarkStart = false;
    	boolean bookmarkEnd = false;
		boolean setted = false;
		int startI = -1;
		int endI = -1;
    	for(int i = 0 ; i < theList.size() ; i++){
    		if(theList.get(i) instanceof JAXBElement ){
    			JAXBElement jax1 = (JAXBElement) theList.get(i);
    			if(jax1.getValue() instanceof CTBookmark){
    				CTBookmark ct = (CTBookmark) jax1.getValue();
    				if(ct.getName().equals(bookmarkName)){
            				bookmarkStart = true;
            				bookmarkEnd = false;
            				startI = i;
    				}
    			} 
    			else if(bookmarkStart && jax1.getValue() instanceof CTMarkupRange){
    				bookmarkStart = false;
        			bookmarkEnd = true;
        			endI = i;
				}

    		}
    		else if(theList.get(i) instanceof R ){
    			R run = (R) theList.get(i);
    			if(run.getContent().size() > 0){
    				for(int j = 0 ; j < run.getContent().size() ; j++ ){
        				if(bookmarkStart && !bookmarkEnd){
        					JAXBElement jax = (JAXBElement) run.getContent().get(j);
            				if(jax.getValue() instanceof Text){
            					if(!setted){
            						if(textToAdd instanceof String){
            							((Text) jax.getValue()).setValue((String)textToAdd);
                					}
                					setted = true;
            					}
            					else{
            						run.getContent().remove(j);
            					}
            				}
        				}
    				}	
    			}
    		}
    		
    	}
    	if(forceSet){
	    	if(!setted){
	    		if(textToAdd instanceof String){
	        		R run = factory.createR();
	                Text t2 = factory.createText();
	                run.getContent().add(t2);   
	                t2.setValue((String)textToAdd);
	                theList.add(endI, run);      
	                endI = -1;
	    		}
	    	}
    	}
    	if(startI >= 0 && endI >=0){
    		theList.remove(endI);
    		theList.remove(startI);
    	}
    	else if(startI >= 0){
    		theList.remove(startI);
    	}    	

	}
	
	private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {

        List<Object> result = new ArrayList<Object>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch))
            result.add(obj); 
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }

        return result; 
    }	
}
