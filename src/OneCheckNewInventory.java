import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Find product ids that are new.
 * Directions - replace CooperClassicsInventory.xls with the latest inventory sheet, and products_export_1.xls with the latest shopify export.
 * Tip - Make sure MPNs in Variant SKU column are not invalid strings like n/a.
 */
public class OneCheckNewInventory {

    public static void main(String[] args) throws Exception {
        int c = 0;
        ArrayList<String> cAndM = new ArrayList<String>();
        ArrayList<String> newItems = new ArrayList<String>();
        
        //System.out.println("Hello, World!");

        //obtaining input bytes from a file  
        FileInputStream fis = new FileInputStream(new File("/Users/eputnam/java_projects/loop-files-and-write-links/ReadExcelFile/products_export_1.xlsx"));  
        Workbook wb = new XSSFWorkbook(fis);
        Workbook wb2 = null;

        try {
            Sheet sheet = wb.getSheetAt(0);
            int totalRows = sheet.getPhysicalNumberOfRows();
            //System.out.println("TOTAL ROWS" + totalRows);
           
            for(int x = 1; x<totalRows; x++){
                Row dataRow = sheet.getRow(x); //get row 1 to row n (rows containing data)

                Cell cell1 = dataRow.getCell(0);
                Cell cell2 = dataRow.getCell(1);
                Cell cell3 = dataRow.getCell(2);

                Cell cellMpn = dataRow.getCell(14);

                if(cellMpn != null){
                    try{
                        String value = cellMpn.getRichStringCellValue().getString().replace("'", "");
                        //System.out.println(value);
                        cAndM.add(value);
                        c++;
                    } catch (java.lang.IllegalStateException e) {
                        System.out.println("ERROR" + e.toString() + " [Value]" + cellMpn.getRowIndex());
                    }
                }
            }

        FileInputStream fis2 = new FileInputStream(new File("/Users/eputnam/java_projects/loop-files-and-write-links/ReadExcelFile/CooperClassicsInventory.xlsx"));  
        wb2 = new XSSFWorkbook(fis2);

        Sheet sheet2 = wb2.getSheetAt(0);
        int totalRows2 = sheet2.getPhysicalNumberOfRows();
            //System.out.println("TOTAL ROWS" + totalRows);
           
            for(int x = 1; x<totalRows2; x++){
                Row dataRow = sheet2.getRow(x); //get row 1 to row n (rows containing data)

                Cell cellMpn = dataRow.getCell(0);

                if(cellMpn != null){
                    try{
                        double value = cellMpn.getNumericCellValue();
                        String myValue = String.valueOf(value).replace(".0", "");
                        if(cAndM.contains(myValue)){
                            //System.out.println("We have this item, so delete it [mpn]" + myValue);
                            //deleteRow(myValue);
                        }else{
                            //System.out.println(myValue);
                            //Add to a new items list
                            newItems.add(myValue);
                        }
                    } catch (java.lang.IllegalStateException e) {
                        System.out.println("ERROR " + e.toString());
                    }
                }
            }

            //Delete any rows that arent in the new list.
            deleteNotNewRows(newItems);
            removeBlankRows();

            System.out.println("Complete [count]" + c);  

            

        } finally{
            wb.close();
            wb2.close();
        }
    }

    public static void deleteNotNewRows(ArrayList<String> newItems) throws IOException {
        String excelPath = new String("/Users/eputnam/java_projects/loop-files-and-write-links/ReadExcelFile/cooper.xlsx");
        Workbook wb3 = null;
        FileOutputStream outFile = null;
        FileInputStream inFile = null;
    
        try {
            inFile = new FileInputStream(excelPath);
            wb3 = new XSSFWorkbook(inFile);
            Sheet sheet3 = wb3.getSheetAt(0);
            int totalRows2 = sheet3.getPhysicalNumberOfRows();
            int lastRowNum = sheet3.getLastRowNum();
            System.out.println(totalRows2);
           
            for(int x = 1; x < totalRows2; x++){
                Row dataRow = sheet3.getRow(x); //get row 1 to row n (rows containing data)
                if(dataRow != null){
                    int rowNo = dataRow.getRowNum();
                    
                    Cell cellMpn = dataRow.getCell(13);
                    if (cellMpn != null){
                        String text = cellMpn.getRichStringCellValue().toString();

                        if(newItems.contains(text)){
                            //System.out.println("Item is new, skip deleting it [mpn]" + text + "[rowno]" + rowNo);
                        }else{
                            System.out.println("Not a new item. Attempting to remove [mpn]" + text + "[rowno]" + rowNo);
                            
                            //remove the row
                            
                            /*if (rowNo >= 0 && rowNo < lastRowNum) {
                                //System.out.println("Row Shifted [mpn]" + text + " [rowno]" + rowNo);
                                sheet3.shiftRows(rowNo + 1, lastRowNum, -1);
                            }*/

                            if (rowNo != lastRowNum) {
                                //System.out.println("Here1 [mpn]" + text + "[rowno]" + rowNo);
                                Row removingRow = sheet3.getRow(rowNo);
                                if(removingRow != null) {
                                    //System.out.println("Here2 [mpn]" + text + "[rowno]" + rowNo);
                                    sheet3.removeRow(removingRow);
                                }
                            }
                            
                            //rewrite file
                            outFile = new FileOutputStream(new File(excelPath));
                            wb3.write(outFile);
                            outFile.close();
                        }
                    }else {
                        System.out.println("Cell Null [row]" + x);
                    }
                } else {
                    System.out.println("Row Null [row]" + x);
                }
            }    
        } catch(Exception e) {
            System.out.println("ERROR " + e.toString());
        } finally {
            if(wb3 != null)
                wb3.close();
            if(outFile != null){
                outFile.close();
            }
            if(inFile != null){
                inFile.close();
            }
        }
    }

    public static void removeBlankRows() throws IOException {
        List<String[]> cellValues = ExcelHandler.extractInfo(new File("/Users/eputnam/java_projects/loop-files-and-write-links/ReadExcelFile/cooper.xlsx"));

		cellValues.forEach(c -> System.out.println(c[0] + ", " + c[1] + ", " + c[2] + ", " + c[3] + ", " + c[4] + ", " + c[5] + ", " + c[6] + ", " + c[7] + ", " + c[8] + ", " + c[9] + ", " + c[10] + ", " + c[11] + ", " + c[12] + ", " + c[13] + ", " + c[14]));

		ExcelHandler.writeToExcel(cellValues, new File("/Users/eputnam/java_projects/loop-files-and-write-links/ReadExcelFile/cooper2.xlsx"));
    }

    
}
