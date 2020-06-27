import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

/**
 * Created by Ummulkiram on 24/08/18.
 */

@Component
public class ApachePOIExcelBuilder {

    private HSSFSheet sheet;
    private HSSFWorkbook workbook;

    @Value("${excel.filePath}")
    private String filePath = "";
    private String sheetName;

    @Value("${excel.dateFormat:m/d/yy h:mm}")
    private String dateFormat;

    private static final Logger logger = LoggerFactory.getLogger(ApachePOIExcelBuilder.class);

    /**
     * An implementation of the buildExcel method in the ExcelBuilder interface, this method returns the workbook object
     * and the excel file path.
     * This method can be used if the path of the generated file is required in addition to the workbook object itself.
     *
     * @param workbookDetails This HashMap workbookDetails should contain the below keys and values :
     *                        key="filePath" - Path where the Excel file is to be saved.
     *                        key="Sheet1" or "Sheet2" -  the hashMaps corresponding to every sheet in the excel.
     *                        Every Sheet HashMap will have the following key value pairs.
     *                        key="rowObjects" - List of HashMaps. Each map contains column name as key and the corresponding value for every record.
     *                        key="sheetName" - The desired name of the sheet in the excel workbook.
     *                        key="formatStrings" -- A HashMap of any format strings you want to pass to be used for a certain data type
     *                                            eg: key="dateFormat" - value = "m/d/yy h:mm"
     *                        key="columnList" - LinkedList of columns needed in the Excel.
     *                        Note: The ColumnList is the list of the desired column name in excel in the order of display.
     *                        The key in the HashMap will be the same as the columnNames in the columnList.
     * @return HashMap that contains both the workbook object and the path of the generated excel.
     * key = "workBook" - the wokrbook object
     * key = "workbookPath" - the path of the generated excel file.
     * @throws Exception
     */
    public HashMap<String, Object> buildExcel(Map<String, Object> workbookDetails) throws Exception {

        HashMap<String, Object> generatedWorkbook = new HashMap<String, Object>();

        generatedWorkbook.put("workBook", buildExcelDocument(workbookDetails));
        generatedWorkbook.put("workbookPath", filePath);

        logger.info("Workbook generated ================");

        return generatedWorkbook;
    }


    private HSSFWorkbook buildExcelDocument(Map<String, Object> workbookDetails)
            throws Exception {

        workbook = new HSSFWorkbook();

        if (workbookDetails.get("filePath") != null) {
            filePath = (String) workbookDetails.get("filePath");
        }

        for (String value : workbookDetails.keySet()) {
            if (workbookDetails.get(value) instanceof HashMap) {
                HashMap<String, Object> sheetDetails = (HashMap) workbookDetails.get(value);

                if (sheetDetails.get("sheetName") != null) {
                    sheetName = (String) sheetDetails.get("sheetName");
                }

                sheet = workbook.createSheet(sheetName);

                //building column header
                List<String> list = (List) sheetDetails.get("columnList");
                LinkedList<String> columnLinkedList = new LinkedList<String>(list);

                buildHeader(columnLinkedList);

                // building rows
                List<HashMap<String, Object>> list2 = (List) sheetDetails.get("rowObjects");
                ArrayList<HashMap<String, Object>> rowObject = new ArrayList<HashMap<String, Object>>(list2);

                buildRows(rowObject,
                        columnLinkedList,
                        (HashMap<String, String>) sheetDetails.get("formatStrings"));
            }
        }

        workbook.write(getFilePathInfo(filePath));
        logger.info(workbook.getSheetAt(0).getSheetName());
        logger.info(workbook.getSheetAt(1).getSheetName());
        return workbook;
    }


    private FileOutputStream getFilePathInfo(String filePath) throws Exception {

        File excelFile = new File(filePath);
        logger.info("Excel filepath: " + excelFile.getPath());
        FileOutputStream fos = new FileOutputStream(excelFile);
        return fos;
    }


    private void buildHeader(LinkedList<String> columnList) {

        Row header = sheet.createRow(0);
        int colNum = 0;
        for (String columns : columnList) {
            header.createCell(colNum++).setCellValue(columns);
        }
    }

    private void buildRows(List<HashMap<String, Object>> rowObject, List<String> columnList, HashMap<String, String> formatStrings) throws Exception {

        int rowNum = 1;

        for (HashMap<String, Object> record : rowObject) {

            Row row = sheet.createRow(rowNum++);
            int cellNum = 0;

            for (String columnName : columnList) {

                logger.info("columnName is: " + columnName);
                logger.info("Column Type is: " + (record.get(columnName) != null ? record.get(columnName).getClass().getName() : ""));
                String colType = "";

                if (record.get(columnName) != null) {
                    colType = record.get(columnName).getClass().getName();
                }
                if (colType.substring(colType.lastIndexOf('.') + 1).equalsIgnoreCase("Date")) {
                    Cell cell = row.createCell(cellNum++);
                    cell.setCellValue((Date) record.get(columnName));
                    cell.setCellStyle(createStyleForDateCells(formatStrings));
                } else if (colType.substring(colType.lastIndexOf('.') + 1).equalsIgnoreCase("Double")) {
                    row.createCell(cellNum++).setCellValue((Double) record.get(columnName));
                } else if (colType.substring(colType.lastIndexOf('.') + 1).equalsIgnoreCase("Long")) {
                    row.createCell(cellNum++).setCellValue((Long) record.get(columnName));
                } else if (colType.substring(colType.lastIndexOf('.') + 1).equalsIgnoreCase("Boolean")) {
                    row.createCell(cellNum++).setCellValue((Boolean) record.get(columnName));
                } else if (colType.substring(colType.lastIndexOf('.') + 1).equalsIgnoreCase("Integer")) {
                    row.createCell(cellNum++).setCellValue((Integer) record.get(columnName));
                } else {
                    row.createCell(cellNum++).setCellValue((String) record.get(columnName));
                }
            }
        }
    }


    /**
     * Defines the excel styles, for ex. the cell borders, text alignment, etc
     */
    public void buildExcelStyle() {

    }


    /**
     * Defines the format of Date cells in the excel file. The desired date format is passed in as the argument.
     *
     * @param formatStrings Map containing the format strings
     *                      Key = "dateFormat" - Values against this key represents the format to be used for date fields.
     * @return object of type CellStyle
     */

    public CellStyle createStyleForDateCells(HashMap<String, String> formatStrings) {

        if (formatStrings != null && formatStrings.get("dateFormat") != null) {
            dateFormat = formatStrings.get("dateFormat");
            logger.info("Date format: " + dateFormat);
        }
        CellStyle cellStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        cellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat(dateFormat));

        return cellStyle;

    }


    /**
     * Turns a text to initcap format
     *
     * @param text
     * @return text in initcap format, for ex. 'text' will change to 'Text'
     */

    public String capitalize(final String text) {
        return Character.toUpperCase(text.charAt(0)) + text.substring(1);
    }

}





