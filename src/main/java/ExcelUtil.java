import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.google.gson.JsonObject;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
    singletone class
 */
public class ExcelUtil {
    private static String address;
    private static ExcelUtil sync=new ExcelUtil("aa");
    private static volatile ExcelUtil instance;
    private ExcelUtil(String address){
        this.address =  address;
    }
    public static ExcelUtil getExcelUtil(String address){
        if(instance == null){
            synchronized (sync){
                if(instance ==null){
                    return new ExcelUtil(address);
                }else return instance;
            }
        }else{
            return instance;
        }
    }

     /*excel cell should not be empty otherwise iterator will skip it.*/
    public  void writeOnExcel(){ // todo : write a correct append to excel code
        XSSFWorkbook workbook= new XSSFWorkbook();
        workbook.createSheet("orders");
        XSSFSheet sheet= workbook.getSheet("orders");
        XSSFTable table= sheet.createTable(null);
        CTTable ctTable= table.getCTTable();
        ctTable.setDisplayName("orders");
        ////////////////////////
        for (int r = 0; r < 2; r++) {
            XSSFRow row = sheet.createRow(r);
            for(int c = 0; c < 3; c++) {
                XSSFCell cell = row.createCell(c);
                if(r == 0) { //first row is for column headers
                    cell.setCellValue("Column"+ (c+1)); //content **must** be here for table column names
                } else {
                    //cell.setCellValue("Data R"+ (r+1) + "C" + (c+1));
                }
            }
        }
        try (FileOutputStream outputStream= new FileOutputStream("D:\\java\\1.gitHub\\labshop\\labShop\\src\\main\\resources\\labshop.xlsx")){
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        /////////////////////////
    }
    /** returns all the tables inside this sheet in an map */
    public Map<String,ArrayNode> convertTablesToListOfArrayNodes(String sheetName){
        XSSFWorkbook workbook=openWorkbook();
        XSSFSheet sheet=workbook.getSheet(sheetName); //1
        List<XSSFTable> tables=sheet.getTables();
        Map<String, List<List<String>>> tablesInMap= new HashMap<>();

        tables.forEach(tbl->{
            int startRow=tbl.getStartRowIndex();
            int endRow = tbl.getEndRowIndex();
            int startCol= tbl.getStartColIndex();
            int endCol= tbl.getEndColIndex();
            List<List<String>> tableInlist= new ArrayList<>();
            for (int row = startRow; row <= endRow ; row++) {
                List<String> rowInList= new ArrayList<>();
                for (int col = startCol; col <=endCol ; col++) {
                    XSSFCell cell=sheet.getRow(row).getCell(col);
                    switch (cell.getCellType()){
                        case STRING:
                            rowInList.add(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            rowInList.add(Double.toString(cell.getNumericCellValue()));
                            break;
                        case BOOLEAN: rowInList.add(Boolean.toString(cell.getBooleanCellValue()));
                            break;
                        case FORMULA: rowInList.add(cell.getStringCellValue());
                            break;
                        default: rowInList.add("2564");
                    }
                }
                tableInlist.add(rowInList);
            }
            tablesInMap.put(tbl.getName(), tableInlist);
        });

        Map<String ,ArrayNode> jsonTables=new HashMap<>();
        tablesInMap.forEach((k,v) -> jsonTables.put(k, toArrayNode(v)));
       /* try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }*/
        return jsonTables;
    }

    private XSSFWorkbook openWorkbook()  {
        XSSFWorkbook workbook=null;
        ClassLoader classLoader=getClass().getClassLoader();

        String excelAddress=this.address;
        try {
            FileInputStream fileInputStream=new FileInputStream(new File(excelAddress));
            workbook=new XSSFWorkbook( fileInputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }

    private ArrayNode toArrayNode(List<List<String>> list){
        JsonObject allObjects=new JsonObject();
        ObjectMapper mapper=new ObjectMapper();
        ArrayNode allNodes= mapper.createArrayNode();
        for(int i=1; i<list.size(); i++){
            JsonObject rowObject= new JsonObject();
            ObjectNode node= mapper.createObjectNode();
            for(int j=0; j< list.get(0).size(); j++){
                rowObject.addProperty(list.get(0).get(j), list.get(i).get(j));
                node.put(list.get(0).get(j), list.get(i).get(j));
            }
            allObjects.add(i+"", rowObject);
            allNodes.add(node);
        }
        //System.out.println(allNodes);
        return allNodes;
    }
}
