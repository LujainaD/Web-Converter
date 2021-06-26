/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package converter;

import au.com.bytecode.opencsv.CSVWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author lujaina
 */
public class Converter {

    public static void main(String[] args) throws Exception {
        //excel file called new.xlsx
        File excelFile = new File("C:\\Users\\lujai\\OneDrive\\Documents\\new.xlsx");
        // 0 is number of the first sheet
        int sheetId = 0;
        convertXLSXFileToCSV(excelFile, sheetId);

    }

    private static void convertXLSXFileToCSV(File excelFile, int sheetId) throws Exception {

        FileInputStream fileStream = new FileInputStream(excelFile);

        // Open the xlsx 
        XSSFWorkbook workBook = new XSSFWorkbook(fileStream);
        //request sheet from the workbook
        XSSFSheet selectedSheet = workBook.getSheetAt(sheetId);
        // number of columns of the head row 
        int numOfRows = selectedSheet.getRow(0).getPhysicalNumberOfCells();

        // Iterate through all the rows in the selected sheet
        Iterator<Row> rowIterator = selectedSheet.iterator();

        //build CSV file
        FileWriter csvFile = new FileWriter("CSVFile.csv");
        //to write in CSV file
        CSVWriter csvOutput = new CSVWriter(csvFile);
        //Loop through rows.
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            int i = 0;//String array
            //the length of your sheet
            String[] csvdata = new String[numOfRows];
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next(); //Fetch CELL
                switch (cell.getCellTypeEnum()) { //Identify CELL type

                    case STRING: //field that represents string cell type 
                        csvdata[i] = cell.getStringCellValue(); 
                        break;
                    case NUMERIC: //field that represents Number cell type 
                        csvdata[i] = cell.toString(); 
                        break;
                }
                i = i + 1;
            }
            csvOutput.writeNext(csvdata);
        }
        csvOutput.close(); //close the CSV file
        workBook.close();// close xlsx file

    }

}
