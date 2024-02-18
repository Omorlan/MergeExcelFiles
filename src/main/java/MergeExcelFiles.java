import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;


public class MergeExcelFiles {

    public static void main(String[] args) {
        String inputFolder = "path"; //change to your path to the folder with xlsx files
        String outputFileName = "combined_data.xlsx"; //output file name

        try {
            List<File> files = listFilesInFolder(inputFolder);
            Workbook combinedWorkbook = new XSSFWorkbook();

            for (File file : files) {
                if (file.getName().endsWith(".xlsx")) {
                    Workbook workbook = new XSSFWorkbook(new FileInputStream(file));
                    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                        Sheet sheet = workbook.getSheetAt(i);
                        Sheet combinedSheet = combinedWorkbook.createSheet(file.getName());
                        copySheet(sheet, combinedSheet);
                    }
                    workbook.close();
                }
            }

            FileOutputStream outputStream = new FileOutputStream(outputFileName);
            combinedWorkbook.write(outputStream);
            combinedWorkbook.close();
            outputStream.close();

            System.out.println("The merged file has been created successfully: " + outputFileName);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<File> listFilesInFolder(String folderPath) {
        File folder = new File(folderPath);
        List<File> fileList = new ArrayList<>();

        if (folder.isDirectory()) {
            for (File file : Objects.requireNonNull(folder.listFiles())) {
                if (file.isFile()) {
                    fileList.add(file);
                }
            }
        }

        return fileList;
    }

    private static void copySheet(Sheet source, Sheet destination) {
        for (int i = 0; i <= source.getLastRowNum(); i++) {
            Row newRow = destination.createRow(i);
            Row row = source.getRow(i);

            if (row != null) {
                newRow.setHeight(row.getHeight());

                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell sourceCell = row.getCell(j);
                    Cell newCell = newRow.createCell(j);

                    if (sourceCell != null) {
                        switch (sourceCell.getCellType()) {
                            case NUMERIC:
                                newCell.setCellValue(sourceCell.getNumericCellValue());
                                break;
                            case STRING:
                                newCell.setCellValue(sourceCell.getStringCellValue());
                                break;
                            case BOOLEAN:
                                newCell.setCellValue(sourceCell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                newCell.setCellFormula(sourceCell.getCellFormula());
                                break;
                            default:
                                break;
                        }
                        CellStyle sourceCellStyle = sourceCell.getCellStyle();
                        CellStyle newCellStyle = destination.getWorkbook().createCellStyle();
                        newCellStyle.cloneStyleFrom(sourceCellStyle);
                        newCell.setCellStyle(newCellStyle);
                    }
                }
            }
        }

        for (int i = 0; i < source.getRow(0).getLastCellNum(); i++) {
            destination.setColumnWidth(i, source.getColumnWidth(i));
        }

        for (int i = 0; i < source.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = source.getMergedRegion(i);
            destination.addMergedRegion(mergedRegion);
        }
    }
}