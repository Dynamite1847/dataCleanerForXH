import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class DataCleaner {
    public static void main(String[] args) throws IOException {
        try (InputStream inp = new FileInputStream("src/main/resources/refdata1.xlsx")) {

            Workbook wb = WorkbookFactory.create(inp);
            cleanData(wb);

            // Write the output to a file
            try (OutputStream fileOut = new FileOutputStream("src/main/resources/refdata2.xlsx")) {
                wb.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static void cleanData(Workbook wb) {
        int valuePosition = 0;
        int reversalFlagPosition = 1;
        int sameValuePosition = 2;
        int reversalRowPosition = 3;
        Sheet sheet = wb.getSheetAt(0);
        for (Row row : sheet) {
            Cell cell = row.getCell(valuePosition, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            double value = cell.getNumericCellValue();
            List<Row> reversalList = new ArrayList<>();
            List<Row> sameValueList = new ArrayList<>();
            for (Row rowInside : sheet) {
                Cell cellValue = rowInside.getCell(valuePosition, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                double valueInside = cellValue.getNumericCellValue();
                if (value == -valueInside) {
                    reversalList.add(rowInside);
                } else if (value == valueInside) {
                    sameValueList.add(rowInside);
                }
            }
            String reversalRowNumbers = new String();
            String sameValueRowNumbers = new String();
            String reversalFlag = new String();
            for (Row reversalRow : reversalList) {
                reversalRowNumbers += String.valueOf(reversalRow.getRowNum() + 1);
                reversalRowNumbers += " , ";
                reversalFlag = "Y";

            }
//            set flag for reversal
            Cell cellReverasalFlag = row.getCell(reversalFlagPosition);
            cellReverasalFlag = createCellIfNull(reversalFlagPosition, row, cellReverasalFlag);
            if (cellReverasalFlag.getStringCellValue().equals(""))
                cellReverasalFlag.setCellValue(reversalFlag);
//                set row numbers for reversal
            Cell cellReverasalPosition = row.getCell(reversalRowPosition);
            cellReverasalPosition = createCellIfNull(reversalRowPosition, row, cellReverasalPosition);
            if (cellReverasalPosition.getStringCellValue().equals(""))
                cellReverasalPosition.setCellValue(reversalRowNumbers);

            for (Row sameValueRow : sameValueList) {
                sameValueRowNumbers += String.valueOf(sameValueRow.getRowNum() + 1);
                sameValueRowNumbers += " ";
            }
            //                set row numbers for samevalue
            Cell cellSameValuePosition = row.getCell(sameValuePosition);
            cellSameValuePosition = createCellIfNull(sameValuePosition, row, cellSameValuePosition);
            if (cellSameValuePosition.getStringCellValue().equals(""))
                cellSameValuePosition.setCellValue(sameValueRowNumbers);
        }
    }

    private static Cell createCellIfNull(int cellPosition, Row row, Cell cell) {
        if (cell == null) {
            cell = row.createCell(cellPosition);
        }
        return cell;
    }
}