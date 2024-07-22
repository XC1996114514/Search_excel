import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class clean_excel_doc {

    public static void main(String[] args) {
        String folderPath = "excel"; // Update with the path to your folder
        List<String> passwordProtectedFiles = new ArrayList<>();

        try {
            scanAndProcessExcelFiles(folderPath, passwordProtectedFiles);
            System.out.println("Processed all Excel files successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }

        if (!passwordProtectedFiles.isEmpty()) {
            System.out.println("The following files require a password to open:");
            for (String fileName : passwordProtectedFiles) {
                System.out.println(fileName);
            }
        }
    }

    public static void scanAndProcessExcelFiles(String folderPath, List<String> passwordProtectedFiles) throws IOException {
        File folder = new File(folderPath);
        File[] files = folder.listFiles((dir, name) -> name.endsWith(".xlsx") || name.endsWith(".xls"));

        if (files != null) {
            for (File file : files) {
                if (!file.canWrite()) {
                    System.out.println("File is read-only: " + file.getName());
                    continue;
                }
                try {
                    Workbook workbook;
                    try (FileInputStream fis = new FileInputStream(file)) {
                        if (file.getName().endsWith(".xlsx")) {
                            workbook = new XSSFWorkbook(fis);
                        } else {
                            workbook = new HSSFWorkbook(fis);
                        }
                    } catch (OLE2NotOfficeXmlFileException e) {
                        System.out.println("File format not supported by this parser, please check if it has password: " + file.getName());
                        continue;
                    } catch (EncryptedDocumentException e) {
                        String password = extractPasswordFromFileName(file.getName());
                        if (password != null) {
                            try (FileInputStream fis = new FileInputStream(file)) {
                                workbook = WorkbookFactory.create(fis, password);
                            } catch (IOException | EncryptedDocumentException ex) {
                                System.out.println("Failed to open password-protected file: " + file.getName());
                                passwordProtectedFiles.add(file.getName());
                                ex.printStackTrace();
                                continue;
                            }
                        } else {
                            System.out.println("Password not found for file: " + file.getName());
                            passwordProtectedFiles.add(file.getName());
                            continue;
                        }
                    }

                    boolean isProtected = false;

                    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                        Sheet sheet = workbook.getSheetAt(i);
                        if (sheet.getProtect()) {
                            isProtected = true;
                            break;
                        }
                    }

                    if (isProtected) {
                        String newFileName = file.getName().replace(".xlsx", "_clear.xlsx").replace(".xls", "_clear.xls");
                        String newFilePath = folderPath + File.separator + newFileName;

                        copyAndUnprotectWorkbook(file, newFilePath, workbook);

                        if (file.delete()) {
                            System.out.println("Copied and unprotected: " + file.getName() + " and deleted original file.");
                        } else {
                            System.out.println("Copied and unprotected: " + file.getName() + " but failed to delete original file.");
                        }
                    }
                } catch (IOException e) {
                    System.out.println("Failed to process file: " + file.getName());
                    e.printStackTrace();
                }
            }
        }
    }

    private static String extractPasswordFromFileName(String fileName) {
        int dotIndex = fileName.indexOf('.');
        if (dotIndex != -1) {
            return fileName.substring(0, dotIndex);
        }
        return null;
    }

    public static void copyAndUnprotectWorkbook(File sourceFile, String destPath, Workbook sourceWorkbook) throws IOException {
        try (Workbook destWorkbook = sourceWorkbook instanceof XSSFWorkbook ? new XSSFWorkbook() : new HSSFWorkbook()) {

            Map<CellStyle, CellStyle> styleMap = new HashMap<>();

            for (int i = 0; i < sourceWorkbook.getNumberOfSheets(); i++) {
                Sheet sourceSheet = sourceWorkbook.getSheetAt(i);
                Sheet destSheet = destWorkbook.createSheet(sourceSheet.getSheetName());

                copySheetContent(sourceSheet, destSheet, styleMap);
            }

            try (FileOutputStream fos = new FileOutputStream(new File(destPath))) {
                destWorkbook.write(fos);
            }
        }
    }

    private static void copySheetContent(Sheet sourceSheet, Sheet destSheet, Map<CellStyle, CellStyle> styleMap) {
        for (int i = 0; i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            Row destRow = destSheet.createRow(i);

            if (sourceRow != null) {
                for (int j = 0; j < sourceRow.getLastCellNum(); j++) {
                    Cell sourceCell = sourceRow.getCell(j);
                    Cell destCell = destRow.createCell(j);

                    if (sourceCell != null) {
                        copyCellContent(sourceCell, destCell, styleMap);
                    }
                }
            }
        }
    }

    private static void copyCellContent(Cell sourceCell, Cell destCell, Map<CellStyle, CellStyle> styleMap) {
        switch (sourceCell.getCellType()) {
            case STRING:
                destCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                destCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                destCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                destCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                destCell.setBlank();
                break;
            default:
                break;
        }
        copyCellStyle(sourceCell, destCell, styleMap);
    }

    private static void copyCellStyle(Cell sourceCell, Cell destCell, Map<CellStyle, CellStyle> styleMap) {
        try {
            CellStyle sourceStyle = sourceCell.getCellStyle();
            CellStyle destStyle = styleMap.get(sourceStyle);

            if (destStyle == null) {
                Workbook destWorkbook = destCell.getSheet().getWorkbook();
                destStyle = destWorkbook.createCellStyle();
                destStyle.cloneStyleFrom(sourceStyle);
                styleMap.put(sourceStyle, destStyle);
            }
            destCell.setCellStyle(destStyle);
        } catch (NullPointerException e) {
            System.out.println("Failed to copy cell style: " + e.getMessage());
        }
    }
}
