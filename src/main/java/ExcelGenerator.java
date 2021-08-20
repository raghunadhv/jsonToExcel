import java.io.File;
import java.io.FileOutputStream;

import java.io.IOException;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;

import models.EntityData;
import models.FieldData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ExcelGenerator {
    private static int rowNo;
    private static int cellNo;
    static void createExcelFile(String directoryPath, List<EntityData> entityDataList) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet("cell types");
        PropertyTemplate pt = new PropertyTemplate();
        createRows(entityDataList, spreadsheet,pt,workbook);
        pt.applyBorders(spreadsheet);
      //Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File("Writesheet.xlsx"));
        workbook.write(out);
        out.close();

    }

    private static void createRows(List<EntityData> entityDataList, XSSFSheet spreadsheet, PropertyTemplate pt, XSSFWorkbook workbook) {
        AtomicBoolean headerPrinted=new AtomicBoolean(false);
        AtomicReference<String> previousEntityName=new AtomicReference<>("");
        CellStyle backgroundStyle = workbook.createCellStyle();
        CellStyle headerBgStyle = workbook.createCellStyle();
        headerBgStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        headerBgStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        entityDataList.forEach(entityData -> {
            cellNo= entityData.getDepth();
            backgroundStyle.setFillForegroundColor((short) (2+entityData.getDepth()));
            backgroundStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            if (!previousEntityName.get().equals(entityData.getEntityName())) {   //skipping header for second row and others
                XSSFRow row = spreadsheet.createRow(rowNo++);
                 XSSFCell cell = row.createCell(cellNo++);
               cell.setCellValue(entityData.getEntityName());
               cell.setCellStyle(backgroundStyle);
            } else {
                cellNo++;
                headerPrinted.set(false);
            }
            createDataTable(spreadsheet, headerPrinted, entityData,pt,backgroundStyle,headerBgStyle);
            headerPrinted.set(false);
            if (!entityData.getEntityDataList().isEmpty()) {
                createRows(entityData.getEntityDataList(),spreadsheet, pt, workbook);
            }
            previousEntityName.set(entityData.getEntityName());
        });


    }

    private static void createDataTable(XSSFSheet spreadsheet, AtomicBoolean headerPrinted, EntityData entityData, PropertyTemplate pt, CellStyle backgroundStyle, CellStyle headerBgStyle) {
        int startRow=rowNo;
        int startCol=cellNo;
        List<FieldData> fieldsList= entityData.getFieldsList();
        if (!fieldsList.isEmpty()) {
            XSSFRow headerRow = null;
            //data header
            if (!headerPrinted.get()) {
                headerRow = spreadsheet.createRow(rowNo++);
            }
            XSSFRow valueRow = spreadsheet.createRow(rowNo++);
            XSSFRow finalHeaderRow = headerRow;
            fieldsList.forEach(fieldData -> {
                if (!headerPrinted.get()) {
                    Cell fieldHeader = finalHeaderRow.createCell(cellNo);
                    fieldHeader.setCellValue(fieldData.getFieldHeaderName());
                    fieldHeader.setCellStyle(headerBgStyle);
                 }
                Cell fieldCell = valueRow.createCell(cellNo);
                fieldCell.setCellValue(fieldData.getFieldName());
                fieldCell.setCellStyle(backgroundStyle);
                cellNo++;
            });
        }
        headerPrinted.set(true);
        if (rowNo>startRow) {

            pt.drawBorders(new CellRangeAddress(startRow, rowNo - 1, startCol, cellNo-1),
                    BorderStyle.MEDIUM, BorderExtent.ALL);
        }
    }

}
