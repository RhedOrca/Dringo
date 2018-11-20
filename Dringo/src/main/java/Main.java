import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;

import java.io.*;
import java.util.*;


public class Main {
    public static void main(String[] args) throws IOException{
        final String excelPath = "/Users/Brandon/Dringo/src/main/java/DringoData.xls";
        Workbook workbook = WorkbookFactory.create(new File(excelPath));
        Sheet sheet = workbook.getSheetAt(0);
        Row[] rows = new Row[10];
        Row row1 = sheet.getRow(0); rows[0] = row1;
        Row row2 = sheet.getRow(1); rows[1] = row2;
        Row row3 = sheet.getRow(2); rows[2] = row3;
        Row row4 = sheet.getRow(3); rows[3] = row4;
        Row row5 = sheet.getRow(4); rows[4] = row5;
        Row row6 = sheet.getRow(5); rows[5] = row6;
        Row row7 = sheet.getRow(6); rows[6] = row7;
        Row row8 = sheet.getRow(7); rows[7] = row8;
        Row row9 = sheet.getRow(8); rows[8] = row9;
        Row row10 = sheet.getRow(9); rows[9] = row10;
        Cell row1Header = row1.getCell(0);
        Cell row2Header = row1.getCell(1);
        Cell row3Header = row1.getCell(2);
        Cell row4Header = row1.getCell(3);
        Cell row5Header = row1.getCell(4);
        Cell row6Header = row1.getCell(5);
        Cell row7Header = row1.getCell(6);
        Cell row8Header = row1.getCell(7);
        Cell row9Header = row1.getCell(8);
        Cell row10Header = row1.getCell(9);
        int lastRow = -1;
        for (int i = 0; i < 10; i++) {
            if (rows[i] != null) {
                if (rows[i].getCell(0).getStringCellValue() != null) {
                    lastRow++; System.out.println(lastRow);
                }
                else {
                    break;
                }
            }
            else {
                break;
            }

        }

        System.out.println("Choose a show:");
        for (int i = 0; i <= lastRow; i++) {
            System.out.println(i+1 + ".) " + rows[i].getCell(0));
        }

        Scanner scanner = new Scanner(System.in);
        int choice = scanner.nextInt();
        int numberOfOptions = sheet.getRow(choice-1).getPhysicalNumberOfCells();

        System.out.println(numberOfOptions);


        String[] options = new String[numberOfOptions];
        for (int i = 0; i < numberOfOptions; i++) {
            options[i] = sheet.getRow(choice-1).getCell(i).getStringCellValue();
            System.out.println(Arrays.toString(options));
        }
        int actualLength = numberOfOptions;
        for (int j = numberOfOptions-1; j > 0; j--) {
            if (options[j] == "") {
                actualLength--;
                continue;
            }
            break;
        }
        String[] optionsFinal = new String[actualLength];
        System.arraycopy(options, 0, optionsFinal, 0, actualLength);
        System.out.println(Arrays.toString(optionsFinal));

        List<String> l = Arrays.asList(optionsFinal);
        Collections.shuffle(l);
        String[] output = l.toArray(new String[actualLength]);

        final Workbook outputFile = new HSSFWorkbook();
        Sheet outputSheet = outputFile.createSheet("DRINGO" + optionsFinal.hashCode());
        CellStyle style = outputSheet.getWorkbook().createCellStyle();
        Font font = outputFile.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 20);
        style.setWrapText(true);
        style.setFont(font);
        Row outputRow1 = outputSheet.createRow(0);
        Row outputRow2 = outputSheet.createRow(1);
        Row outputRow3 = outputSheet.createRow(2);
        Row outputRow4 = outputSheet.createRow(3);
        Row outputRow5 = outputSheet.createRow(4);
        int rowCounter = 0;
        for (Row outputRow : outputSheet) {
            for (int k = 0 + 5*rowCounter; k < 5*rowCounter + 5; k++) {
                outputRow.createCell(k-5*rowCounter).setCellValue(output[k]);
            }
            rowCounter++;
        }
        outputRow3.getCell(2).setCellValue("Free!");

        float height = (float) 110.0;

        outputRow1.setHeightInPoints(height);
        outputRow2.setHeightInPoints(height);
        outputRow3.setHeightInPoints(height);
        outputRow4.setHeightInPoints(height);
        outputRow5.setHeightInPoints(height);
        for (int m = 0; m<5;m++) {
            outputSheet.setColumnWidth(m, 8000);
        }
        for (Row row : outputSheet) {
            for (Cell cell : row) {
                cell.setCellStyle(style);
                CellUtil.setAlignment(cell, HorizontalAlignment.CENTER);
                CellUtil.setVerticalAlignment(cell, VerticalAlignment.CENTER);
            }
        }

        OutputStream fileOut = new FileOutputStream("Dringo.xls");
        outputFile.write(fileOut);

        System.out.println(Arrays.toString(output));
        Random rand = new Random();
        System.out.println(rand.nextInt(6));
        System.out.println(rand.nextInt(17));

    }
}
