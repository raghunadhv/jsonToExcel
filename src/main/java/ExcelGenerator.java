import java.io.File;
import java.io.FileOutputStream;

import java.io.IOException;
import java.util.List;
import java.util.Set;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

import models.EntityData;
import models.FieldData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ExcelGenerator {
    private static int rowNo;
    private static int cellNo;
    static void createExcelFile(String directoryPath, List<EntityData> entityDataList) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        //Create row object

        XSSFSheet spreadsheet = workbook.createSheet("cell types");

        createRows(entityDataList, spreadsheet);


        //Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File("Writesheet.xlsx"));
        workbook.write(out);
        out.close();

    }

    private static void createRows(List<EntityData> entityDataList,  XSSFSheet spreadsheet) {
        AtomicBoolean headerPrinted=new AtomicBoolean(false);
        AtomicReference<String> previousEntityName=new AtomicReference<>("");
        entityDataList.forEach(entityData -> {
            cellNo= entityData.getDepth()-1;
            if (!previousEntityName.get().equals(entityData.getEntityName())) {
                XSSFRow row = spreadsheet.createRow(rowNo++);

                XSSFCell cell = row.createCell(cellNo++);

                //row = spreadsheet.createRow(rowNo.getAndIncrement());

                cell.setCellValue(entityData.getEntityName());
            } else {
               // rowNo--;
                cellNo++;
            }
            List<FieldData> fieldsList= entityData.getFieldsList();
            if (!fieldsList.isEmpty()) {
                XSSFRow headerRow = null;
                if (!headerPrinted.get()) {
                    headerRow = spreadsheet.createRow(rowNo++);
                }
                XSSFRow valueRow = spreadsheet.createRow(rowNo++);
                XSSFRow finalHeaderRow = headerRow;
                fieldsList.forEach(fieldData -> {
                    if (!headerPrinted.get()) {
                        Cell fieldHeader = finalHeaderRow.createCell(cellNo);
                        fieldHeader.setCellValue(fieldData.getFieldHeaderName());
                     }
                    Cell fieldCell = valueRow.createCell(cellNo);
                    fieldCell.setCellValue(fieldData.getFieldName());
                    cellNo++;
                });
            }
            headerPrinted.set(true);
            if (!entityData.getEntityDataList().isEmpty()) {
                createRows(entityData.getEntityDataList(),spreadsheet);
            }
            previousEntityName.set(entityData.getEntityName());
        });


    }

    public static void main(String args[]) throws IOException {
      //  createExcelFile("");
    }
}
