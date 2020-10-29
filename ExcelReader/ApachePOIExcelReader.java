package com.cdplatform.document.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;

@Slf4j
@Component("ApachePOIReader")
public class ApachePOIExcelReader implements ExcelReader {

    private byte[] document;

    @Override
    public HashMap<String, List<Object>> readExcelFromPath(HashMap<String, String> recordTypes, String fileLocation) throws Exception {

        String fileFormat = getFileFormat(fileLocation);
        FileInputStream stream = new FileInputStream(new File(fileLocation));
        Workbook workbook = getExcelWorkbook(stream, fileFormat);
        List<Sheet> listOfSheets = getSheets(workbook);

        return readSheets(listOfSheets, recordTypes);

    }


    @Override
    public HashMap<String, List<HashMap<String, Object>>> readExcelFromPathGeneric(String fileLocation) throws Exception {

        String fileFormat = getFileFormat(fileLocation);
        FileInputStream stream = new FileInputStream(new File(fileLocation));
        Workbook workbook = getExcelWorkbook(stream, fileFormat);
        List<Sheet> listOfSheets = getSheets(workbook);

        return readSheetsGeneric(listOfSheets);

    }


    @Override
    public HashMap<String, List<Object>> readExcelFromURL(HashMap<String, String> recordTypes, String fileLocation) throws Exception {

        String fileFormat = getFileFormat(fileLocation);
        final URL fileUrl = new URL(fileLocation);
        InputStream stream = fileUrl.openStream();
        Workbook workbook = getExcelWorkbook(stream, fileFormat);
        List<Sheet> listOfSheets = getSheets(workbook);

        return readSheets(listOfSheets, recordTypes);

    }


    @Override
    public HashMap<String, List<HashMap<String, Object>>> readExcelFromUrlGeneric(String fileLocation) throws Exception {

        String fileFormat = getFileFormat(fileLocation);
        final URL fileUrl = new URL(fileLocation);
        InputStream stream = fileUrl.openStream();
        Workbook workbook = getExcelWorkbook(stream, fileFormat);
        List<Sheet> listOfSheets = getSheets(workbook);

        return readSheetsGeneric(listOfSheets);

    }


    public Workbook getExcelWorkbook(InputStream stream, String fileFormat) throws Exception {

        // ByteArrayInputStream file = new ByteArrayInputStream(document); // remove document after test

        if (fileFormat.equalsIgnoreCase("xlsx")) {

            return new XSSFWorkbook(stream);
        } else if (fileFormat.equalsIgnoreCase("xls")) {

            return new HSSFWorkbook(stream);
        }
        return null;
    }


    private String getFileFormat(String fileLocation) {

        String fileFormat = fileLocation.substring(fileLocation.lastIndexOf(".") + 1);

        return fileFormat;

    }


    private List<Sheet> getSheets(Workbook workbook) {

        int numOfSheets = workbook.getNumberOfSheets();
        List<Sheet> listOfSheets = new ArrayList<Sheet>();

        for (int i = 0; i < numOfSheets; i++) {
            listOfSheets.add(workbook.getSheetAt(i));
        }

        return listOfSheets;
    }


    private HashMap<String, List<Object>> readSheets(List<Sheet> listOfSheets, HashMap<String, String> recordTypes) throws Exception {

        HashMap<String, List<Object>> excelContent = new HashMap<String, List<Object>>();

        for (Sheet sheet : listOfSheets) {

            LinkedList<String> headers = getHeaders(sheet);
            excelContent.put(sheet.getSheetName(), constructRowObjects(sheet, headers, recordTypes.get(sheet.getSheetName())));

        }

        return excelContent;
    }

    private HashMap<String, List<HashMap<String, Object>>> readSheetsGeneric(List<Sheet> listOfSheets) throws Exception {

        HashMap<String, List<HashMap<String, Object>>> excelContent = new HashMap<String, List<HashMap<String, Object>>>();

        for (Sheet sheet : listOfSheets) {

            LinkedList<String> headers = getHeaders(sheet);
            excelContent.put(sheet.getSheetName(), constructRowGeneric(sheet, headers));

        }

        return excelContent;
    }


    private LinkedList<String> getHeaders(Sheet sheet) {

        LinkedList<String> header = new LinkedList<String>();
        Row headerRow = sheet.getRow(0);

        for (Cell cell : headerRow) {
            switch (cell.getCellTypeEnum()) {
                case STRING:
                    header.add(cell.getStringCellValue());
                    break;
            }
        }
        return header;
    }


    private List<Object> constructRowObjects(Sheet sheet, LinkedList<String> headers, String recordType) throws Exception {

        List<Object> sheetContent = new ArrayList<Object>();

        for (int rowStart = 1; rowStart <= sheet.getPhysicalNumberOfRows() - 1; rowStart++) {

            LinkedList<Object> rowLine = extractRowValues(sheet, rowStart);
            sheetContent.add(constructObjects(headers, rowLine, recordType));

        }
        return sheetContent;
    }


    private List<HashMap<String, Object>> constructRowGeneric(Sheet sheet, LinkedList<String> headers) throws Exception {

        List<HashMap<String, Object>> sheetContent = new ArrayList<HashMap<String, Object>>();

        for (int rowStart = 1; rowStart <= sheet.getPhysicalNumberOfRows() - 1; rowStart++) {

            LinkedList<Object> rowLine = extractRowValues(sheet, rowStart);
            sheetContent.add(constructRowMaps(headers, rowLine));

        }
        return sheetContent;
    }


    private LinkedList<Object> extractRowValues(Sheet sheet, int rowStart) {

        LinkedList<Object> rowLine = new LinkedList<Object>();

        Row row = sheet.getRow(rowStart);
        for (Cell cell : row) {
            switch (cell.getCellTypeEnum()) {
                case STRING:
                    rowLine.add(cell.getStringCellValue());
                    break;
                case NUMERIC:
                    rowLine.add(cell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    rowLine.add(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    rowLine.add(cell.getCellFormula());
                    break;
                default:
            }
        }

        return rowLine;
    }


    private Object constructObjects(LinkedList<String> headers, LinkedList<Object> rowLine, String recordType) throws Exception {

        Class<?> recordClass = Class.forName(recordType);
        Object recordObject = recordClass.newInstance();

        for (int i = 0; i < rowLine.size(); i++) {
            Field f1 = recordObject.getClass().getDeclaredField(headers.get(i));
            f1.setAccessible(true);
            f1.set(recordObject, rowLine.get(i));
        }

        return recordObject;

    }


    private HashMap<String, Object> constructRowMaps(LinkedList<String> headers, LinkedList<Object> rowLine) throws Exception {

        HashMap<String, Object> row = new HashMap<String, Object>();
        for (int i = 0; i < rowLine.size(); i++) {

            row.put(headers.get(i), rowLine.get(i));
        }

        return row;

    }


}
