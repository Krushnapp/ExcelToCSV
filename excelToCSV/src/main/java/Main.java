import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main {


    static void xlsx(File inputFile, File outputFile) {
        // For storing data into CSV files
        StringBuffer data = new StringBuffer();

        try {
            FileOutputStream fos = new FileOutputStream(outputFile);
            // Get the workbook object for XLSX file
            FileInputStream fis = new FileInputStream(inputFile);
            Workbook workbook = new XSSFWorkbook();

            String ext = FilenameUtils.getExtension(inputFile.toString());

            if (ext.equalsIgnoreCase("xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (ext.equalsIgnoreCase("xls")) {
                workbook = new HSSFWorkbook(fis);
            }

            // Get first sheet from the workbook

            int numberOfSheets = workbook.getNumberOfSheets();
            Row row;
            Cell cell;
            // Iterate through each rows from first sheet

            Boolean isHead = true;

            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();



                while (rowIterator.hasNext()) {
                    row = rowIterator.next();
                    // For each row, iterate through each columns
                    Iterator<Cell> cellIterator = row.cellIterator();
                    int cellCount = 0;
                    while (cellIterator.hasNext()) {
                        String discp = "";

                        cell = cellIterator.next();

                        switch (cell.getCellType()) {
                            case BOOLEAN:
                                data.append(cell.getBooleanCellValue() + ",");

                                break;
                            case NUMERIC:
                                if(cellCount == 4 && !isHead ){
                                    double val = cell.getNumericCellValue();
                                    double nVal = val + 0.2*val;

                                    cell.setCellValue(nVal);

                                }
                                data.append(cell.getNumericCellValue() + ",");

                                break;
                            case STRING:
                                if(cell.getStringCellValue().equalsIgnoreCase("MfrPN"))
                                    cell.setCellValue("Mfr P/N");
                                if(cell.getStringCellValue().equalsIgnoreCase("Cost"))
                                    cell.setCellValue("Price");
                                if(cell.getStringCellValue().equalsIgnoreCase("Coo"))
                                    cell.setCellValue("COO");
                                if(cellCount == 6 && !isHead && !cell.getStringCellValue().isBlank() && cell.getStringCellValue().length()>300)
                                    discp = cell.getStringCellValue().substring(0,300);
                                    cell.setCellValue(discp);

                                if(cell.getStringCellValue().isBlank() && !isHead )
                                    if(cellCount == 5)
                                        cell.setCellValue("TW");
                                    if(cellCount == 8)
                                        cell.setCellValue("EA");
                                data.append(cell.getStringCellValue() + ",");
                                break;

                            case BLANK:
                                data.append("" + ",");
                                break;
                            default:
                                data.append(cell + ",");

                        }

                        cellCount++;
                    }

                    if(isHead)
                        isHead = false;
                    data.append('\n'); // appending new line after each row
                }

            }
            fos.write(data.toString().getBytes());
            fos.close();

        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }

    // testing the application

    public static void main(String[] args) {
        // int i=0;
        // reading file from desktop


        /*
        * Need to change the input and output file location before executing the programme.
        * */

        File inputFile = new File("C:\\Users\\shree\\Desktop\\napcloud\\InputXLS.xlsx"); //provide your path
        // writing excel data to csv
        File outputFile = new File("C:\\Users\\shree\\Desktop\\napcloud\\outputCSV123");  //provide your path
        if(outputFile.exists()){
            outputFile.delete();
        }

        xlsx(inputFile, outputFile);
        System.out.println("Conversion of " + inputFile + " to flat file: "
                + outputFile + " is completed");
    }
}
