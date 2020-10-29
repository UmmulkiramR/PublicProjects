package com.cdplatform.document.utils;

import org.springframework.beans.factory.annotation.Autowired;

import java.util.HashMap;
import java.util.List;

public class TestFileUpload {

    //static String fileLocation = "http://localhost:8086/docs/get/Client_Batch_Upload1581577097846.xlsx";
    static String fileLocation = "/Users/ummulkiram 1/Documents/Client_Batch_Upload.xlsx";

    /*@Autowired
    static ApachePOIExcelReader apachePOIExcelReader;*/


    public static void main(String args[]) throws Exception{

        ApachePOIExcelReader apachePOIExcelReader = new ApachePOIExcelReader();


        HashMap<String,String> recordTypes = new HashMap<>();
        recordTypes.put("clientUpdate","com.cdplatform.document.vo.ClientBatchUpdate");
        recordTypes.put("issuerUpdate","com.cdplatform.document.vo.IssuerBatchUpdate");

        HashMap<String, List<Object>> excelContents =  apachePOIExcelReader.readExcelFromPath(recordTypes, fileLocation);

        System.out.println("excelContents: "+excelContents);

    }
}
