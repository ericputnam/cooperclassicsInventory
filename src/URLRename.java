import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;

import org.apache.commons.text.WordUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 

/**
 * Class that takes the cooper.xls and reformats image URLs into the shopify specific format, and also pulls down a copy of the image. After this program completes '
 * you'll find a copy of the image at /Users/eputnam/cooper/. From there you can bulk upload these images at ...
 * Directions - Make sure to update the latest version of cooper2.xslx from OneCheckNewInventory
 */
public class URLRename {
    public static void main(String[] args) throws Exception {
        
        //System.out.println("Hello, World!");

        //obtaining input bytes from a file  
        FileInputStream fis = new FileInputStream(new File("/Users/eputnam/java_projects/loop-files-and-write-links/ReadExcelFile/cooper2.xlsx"));  
        Workbook wb = new XSSFWorkbook(fis);

        try {
            Sheet sheet = wb.getSheetAt(0);
            int totalRows = sheet.getPhysicalNumberOfRows();
            //System.out.println("TOTAL ROWS" + totalRows);
           
            for(int x = 1; x<totalRows; x++){
                Row dataRow = sheet.getRow(x); //get row 1 to row n (rows containing data)

                Cell cell1 = dataRow.getCell(0);
                Cell cell2 = dataRow.getCell(1);
                Cell cell3 = dataRow.getCell(2);

                //System.out.println(cell1.getStringCellValue());  
                

                String url = cell2.getStringCellValue();
                String sourceURL = url;
                //System.out.println(url);  

                
                // Match regex against input
                //Matcher matcher = pattern.matcher(url);
                // Use results...
                //System.out.println("Matched" + matcher.matches());
                String url2 = url.replace("https://i0.wp.com/cooperclassics.com/wp-content/uploads/", "");
                //String[] tokens= url2.split("//");
                String[] arrOfStr = url2.split("/", 3);
 
                String a = arrOfStr[arrOfStr.length - 1];
                //System.out.println(a);

                String[] arrFinal = a.split("\\?");
                String aFinal = arrFinal[0];

                String realname = cell3.getStringCellValue();
                String realname1 = realname.replace("https://cooperclassics.com/product/", "");
                String realname2 = realname1.replace("/", "").replace("-", " ");
                String finalname = WordUtils.capitalize(realname2);
                
                //System.out.println(finalname);
                System.out.println("https://cdn.shopify.com/s/files/1/0603/7356/5601/files/" + aFinal + "?v=1642340850");
                //System.out.println("");

                File file = new File("/Users/eputnam/cooper/" + cell1.getStringCellValue());
                if(!file.exists())
                    file.mkdir();
                
                
                /* uncomment */ //saveImage(sourceURL, aFinal, cell1.getStringCellValue());

                /* uncomment */ //TimeUnit.SECONDS.sleep(6);
            }
            
            System.out.println("Complete");  

        } finally{
            wb.close();
        }
    }  

    public static void saveImage(String imageUrl, String destinationFile, String folderName) throws IOException {
        InputStream is = null;
        OutputStream os = null;
    
        try{
            /* uncomment */ //System.out.println("Getting.. " + imageUrl);
            URL url = new URL(imageUrl);
            is = url.openStream();
            os = new FileOutputStream("/Users/eputnam/cooper/" + folderName + "/" + destinationFile);

            byte[] b = new byte[2048];
            int length;
        
            while ((length = is.read(b)) != -1) {
                os.write(b, 0, length);
            }
        } catch (IOException e){
            e.printStackTrace();
        } finally {
            is.close();
            os.close();
        }
    
        
    }
}
