import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class poiDemo {

    public static void main (String[] args){

        //createworkbook("Employees", "Records");
        //readexcel("Employees", "Records");
        //appendrecord("Employees", "Records", "4", "Oreo", "Food Inspector");
    }

    public  static  void appendrecord(String wb, String ws, String id, String name, String department) {

        try {
            FileInputStream file = new FileInputStream(new File (wb + ".xlsx"));

            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet worksheet = workbook.getSheet(ws);

            int rowlastnum = worksheet.getLastRowNum();
            Row newrow = worksheet.createRow(rowlastnum + 1);
            Cell cell1 = newrow.createCell(0);
            cell1.setCellValue(id);

            Cell cell2 = newrow.createCell(1);
            cell2.setCellValue(name);

            Cell cell3 = newrow.createCell(2);
            cell3.setCellValue(department);

            //write to file
            FileOutputStream out = new FileOutputStream(wb + ".xlsx");
            workbook.write(out);
            System.out.println("New row successfully added!");
            out.close();
            file.close();

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
        public static void readexcel(String workbookname, String worksheetname){

        try {
            File file = new File(workbookname + ".xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheet(worksheetname);
            //XSSFSheet sheet1 = workbook.getSheetAt(0);

            //loop over rows in sheet
            Iterator<Row> rowIterator = sheet.rowIterator();
            while(rowIterator.hasNext()){
                Row row = rowIterator.next();

                //loop over columns in each row
                Iterator<Cell> cellIterator = row.cellIterator();
                while(cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    System.out.print(cell.getStringCellValue() + "\t");
                }
                System.out.print("\n");
            }
            System.out.println("------END------");


        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }

    }

    public static void readexcel(){

        try {
            File file = new File("Employees.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheet("Records");
            //XSSFSheet sheet1 = workbook.getSheetAt(0);

            //loop over rows in sheet
            Iterator<Row> rowIterator = sheet.rowIterator();
            while(rowIterator.hasNext()){
                Row row = rowIterator.next();

                //loop over columns in each row
                Iterator<Cell> cellIterator = row.cellIterator();
                while(cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    System.out.print(cell.getStringCellValue() + "\t");
                }
                System.out.print("\n");
            }
            System.out.println("------END------");


        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }

    }

    public static void createworkbook(){

        //write to xlsx
        //create instance workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Employees");
        /*XSSFSheet sheet1 = workbook.createSheet("Employees1");
        XSSFSheet sheet2 = workbook.createSheet("Employees2");
        XSSFSheet sheet3 = workbook.createSheet("Employees3");*/


        //data
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"id", "name", "department"});
        data.put("2", new Object[] {"1", "joseph", "mis"});
        data.put("3", new Object[] {"2", "ryan", "hr"});
        data.put("4", new Object[] {"3", "didi", "accounting"});

        Set<String> keyset = data.keySet();

        int rownum = 0;
        //loop each keyset
        for(String key:keyset){

            Row row = sheet.createRow(rownum++);
            Object[] obj = data.get(key);
            int cellnum = 0;
            //loop each column in earch row
            for(Object o:obj){
                Cell cell = row.createCell(cellnum++);
                cell.setCellValue(o.toString());
            }

        }

        //write file in filesystem
        try {
            FileOutputStream out = new FileOutputStream(new File("Employees.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("Write xlsx ok!");

        }catch (Exception e){
            System.out.println(e);
        }
    }

    public static void createworkbook(String workbookName, String worksheetName){
        //write to xlsx
        //create instance workbook

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(worksheetName);

        //data
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"id", "name", "department"});
        data.put("2", new Object[] {"1", "joseph", "mis"});
        data.put("3", new Object[] {"2", "ryan", "hr"});
        data.put("4", new Object[] {"3", "didi", "accounting"});

        Set<String> keyset = data.keySet();

        int rownum = 0;
        //loop each keyset
        for(String key:keyset){

            Row row = sheet.createRow(rownum++);
            Object[] obj = data.get(key);
            int cellnum = 0;
            //loop each column in earch row
            for(Object o:obj){
                Cell cell = row.createCell(cellnum++);
                cell.setCellValue(o.toString());
            }

        }

        //write file in filesystem
        try {
            FileOutputStream out = new FileOutputStream(new File(workbookName + ".xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("Write xlsx ok!");

        }catch (Exception e){
            System.out.println(e);
        }
    }
}
