package excel.json;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.net.URL;
import java.util.HashMap;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

public class JsonToExcel {

	public static void exportswaggerUrlToExcel(String swaggerUrl) throws Exception{
		// TODO Auto-generated method stub

		XSSFWorkbook workbook = new XSSFWorkbook();
        
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("API Details");

        String jsonString = readUrl(swaggerUrl);
        JSONParser parser = new JSONParser();
        
        JSONObject json = (JSONObject) parser.parse(jsonString);
        
        JSONObject paths = (JSONObject)json.get("paths");
        
        HashMap<String,JSONObject> apiSet = (HashMap<String, JSONObject>) paths;
        int count = 0;
        int rownum = 0;
        
        Row row = sheet.createRow(rownum++);
        String [] objArr = new String[]{"S.No.","API_URL","CONTROLLER", "DESCRIPTION", "RESPONSE_TYPE"};
        int cellnum = 0;
        
        for (String obj : objArr)
        {
           Cell cell = row.createCell(cellnum++);
                cell.setCellValue((String)obj);
            
        }
        
        rownum++;
        
        for (Entry<String, JSONObject> e : apiSet.entrySet()) {
        	
        	String key = e.getKey();
            JSONObject api = e.getValue();
            System.out.println();
            JSONObject method = (JSONObject) api.get("post");
            if(method == null){
            	method = (JSONObject) api.get("get");
            	
            	if(method == null){
            		method = (JSONObject) api.get("delete");
            	}
            }
            
            String controller = ((JSONArray)method.get("tags")).toString();
            String description = (String) method.get("summary");
            
            JSONObject schema = (JSONObject)(((JSONObject)((JSONObject)method.get("responses")).get("200")).get("schema"));
            String responses = null;
            if(schema != null){
            	
            responses = (String)schema.get("$ref");
            
            if(responses == null){
            	
            	if(((String)schema.get("type")).equals("array"))
            		responses = (String)((JSONObject)schema.get("items")).get("$ref");
            	else if(((String)schema.get("type")).equals("object"))
            		responses = (String)((JSONObject)schema.get("additionalProperties")).get("type");
            }
            
            if(responses !=null){
            	String[]  res= responses.split("/");
            	
            	responses = res[res.length - 1];
            }
            }
            System.out.println("api:" + key + "\tcontroller:" + controller + "\tdescription:" + description + "\tresponses:" + responses);
            
            row = sheet.createRow(rownum++);
            objArr = new String[]{key,controller, description, responses};
            cellnum = 0;
            
            Cell cell = row.createCell(cellnum++);
            cell.setCellValue(++count);
            
            for (String obj : objArr)
            {
               cell = row.createCell(cellnum++);
                    cell.setCellValue((String)obj);
                
            }
        
        }
        
        System.out.println(count);
       
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("API_DETAILS.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("API_DETAILS.xlsx written successfully on disk.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
	
	private static String readUrl(String urlString) throws Exception {
	    BufferedReader reader = null;
	    try {
	        URL url = new URL(urlString);
	        reader = new BufferedReader(new InputStreamReader(url.openStream()));
	        StringBuffer buffer = new StringBuffer();
	        int read;
	        char[] chars = new char[1024];
	        while ((read = reader.read(chars)) != -1)
	            buffer.append(chars, 0, read); 

	        return buffer.toString();
	    } finally {
	        if (reader != null)
	            reader.close();
	    }
	}
	

}
