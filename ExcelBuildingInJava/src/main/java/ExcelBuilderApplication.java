
import org.springframework.boot.autoconfigure.SpringBootApplication;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;

@SpringBootApplication
public class ExcelBuilderApplication {

    public static void main(String args[]) throws Exception{

        HashMap<String, Object> sheet1 = createSheet(3);
        HashMap<String, Object> sheet2 = createSheet(5);

        HashMap<String, Object> workbookDetails = new HashMap<>();
        workbookDetails.put("Sheet1", sheet1);
        workbookDetails.put("Sheet2", sheet2);
        workbookDetails.put("filePath", "C:\\Users\\ExcelBuilder\\excel.xls");

        ApachePOIExcelBuilder apachePOIExcelBuilder2 = new ApachePOIExcelBuilder();
        HashMap<String, Object> excel = apachePOIExcelBuilder2.buildExcel(workbookDetails);

    }

    private static HashMap<String, Object> createSheet(int numberOfRows) {
        HashMap<String, Object> sheet1 = new HashMap<String, Object>();
        LinkedList<String> columns = createColumnList(getRowType());
        sheet1.put("columnList", columns);
        sheet1.put("rowObjects", createRows(columns, numberOfRows));
        sheet1.put("sheetName", "Sheet"+numberOfRows);

        return sheet1;
    }

    private static Object getRowType() {
        return null;
    }

    private static List<HashMap<String, Object>> createRows(LinkedList<String> columns, int numberOfRows) {

        ArrayList<HashMap<String, Object>> rowList = new ArrayList<>();

        for (int i = 0; i <= numberOfRows; i++) {
            HashMap<String, Object> row = new HashMap<String, Object>();
            for (String columnName : columns) {
                row.put(columnName, columnName + " value");
            }
            rowList.add(row);
        }

        return rowList;
    }


    private static LinkedList<String> createColumnList(Object rowType) {
        LinkedList<String> columns = new LinkedList<>();
     // the rowType object here is passed as null only for demonstration purpose.
     // for actual applications this argument can be used to create different column list for different sheets.
        if(rowType == null) {
            columns.add("column1");
            columns.add("column2");
            columns.add("column3");
        }
        return columns;
    }

}
