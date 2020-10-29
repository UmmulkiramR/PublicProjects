package com.cdplatform.document.utils;

import java.util.HashMap;
import java.util.List;

public interface ExcelReader {

    public HashMap<String, List<Object>> readExcelFromPath(HashMap<String,String> recordTypes, String fileLocation) throws Exception;

    public HashMap<String, List<Object>> readExcelFromURL(HashMap<String, String> recordTypes, String fileLocation) throws Exception;

    public HashMap<String, List<HashMap<String, Object>>> readExcelFromUrlGeneric(String fileLocation) throws Exception;

    public HashMap<String, List<HashMap<String, Object>>> readExcelFromPathGeneric(String fileLocation) throws Exception;
}
