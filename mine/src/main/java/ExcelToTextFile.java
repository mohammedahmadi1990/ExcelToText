import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.formula.DataValidationEvaluator;
import org.apache.poi.ss.formula.EvaluationConditionalFormatRule;
import org.apache.poi.ss.formula.WorkbookEvaluatorProvider;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class ExcelToTextFile {
    private static final String SOURCE_EXTENSION = "xlsx";
    private static final String TARGET_EXTENSION = "txt";
    private static final Predicate<String> containsTarget = path -> path.contains("\\target");
    private static final Predicate<String> isExcelFile = path -> path.endsWith("xlsx");

    private final String searchDirectory;

    public ExcelToTextFile(String searchDirectory) {
        this.searchDirectory = searchDirectory;
    }

    public void generateTextFilesFromExcelFile() {
        try {
            getAllExcelFiles().forEach(this::convertExcelToTextFile);
        } catch (IOException e) {
            ;
        }
    }

    private void convertExcelToTextFile(String pathToExcel) {
        StringBuilder fileContent = new StringBuilder();
        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(pathToExcel))) {
            File file = new File(createTextPath(pathToExcel));
            for (Sheet sheet : workbook) {
                appendSheetName(sheet, fileContent);
                appendSheetContent(sheet, fileContent);
                fileContent.append(System.lineSeparator());
            }
            Files.write(file.toPath(), String.valueOf(fileContent).getBytes(StandardCharsets.UTF_8));
        } catch (Exception e) {
            ;
        }
    }

    private String createTextPath(String fileName) {
        StringBuilder filePath = new StringBuilder();
        int extensionIndex = fileName.lastIndexOf(SOURCE_EXTENSION);

        if (extensionIndex > -1) {
            filePath.append(fileName, 0, extensionIndex)
                    .append(TARGET_EXTENSION);
        } else {
            filePath.append(fileName);
        }
        return filePath.toString();
    }

    private void appendSheetName(Sheet sheet, StringBuilder fileContent) {
        fileContent.append("============").append(sheet.getSheetName()).append("========================\n");
    }

    public static boolean isMergedRegion(Sheet sheet, int row, int column) {
        final int sheetMergeCount = sheet.getNumMergedRegions();
        CellRangeAddress ca;
        for (int i = 0; i < sheetMergeCount; i++) {
            ca = sheet.getMergedRegion(i);
            if (row >= ca.getFirstRow() && row <= ca.getLastRow() && column >= ca.getFirstColumn() && column <= ca.getLastColumn()) {
                return true;
            }
        }
        return false;
    }

    private void appendSheetContent(Sheet sheet, StringBuilder fileContent) {
        List<List<String>> sheetTable = new LinkedList<>();
        List<Integer> maxColumnLengths = new LinkedList<>();

        for (Row row : sheet) {
            // For some rows, getLastCellNum returns -1. These rows must be igonred
            int step = 0;
            if (row.getLastCellNum() > 0) {
                List<String> columns = new LinkedList<>();

                for (Cell cell : row) {
                    String cellContent;

                    int freeCount = cell.getColumnIndex() - step;

                    if (!CellType._NONE.equals(cell.getCellType()) && !CellType.BLANK.equals(cell.getCellType())) {
                        if(freeCount >0 || isMergedRegion(sheet,row.getRowNum(),step)){ //|| (freeCount==0 && step==0 && cell.getColumnIndex()==0)){
                            cellContent = "";
                            columns.add(cellContent);
                        }
                        cellContent = getCellValueAndCharacteristics(cell, sheet.getWorkbook().getFontAt(cell.getCellStyle().getFontIndexAsInt()));
                    } else {
                        cellContent = "";
                    }
                    step = step+1;

//                    if (!CellType._NONE.equals(cell.getCellType()) && !CellType.BLANK.equals(cell.getCellType())) {
//                        cellContent = getCellValueAndCharacteristics(cell, sheet.getWorkbook().getFontAt(cell.getCellStyle().getFontIndexAsInt()));
//                    } else {
//                        cellContent = "";
//                    }

                    columns.add(cellContent);

                    int cellLength = cellContent.length();

                    if (maxColumnLengths.size() > cell.getColumnIndex()) {
                        maxColumnLengths.set(cell.getColumnIndex(), Math.max(cellLength, maxColumnLengths.get(cell.getColumnIndex())));
                    } else {
                        maxColumnLengths.add(cellLength);
                    }
                }

                sheetTable.add(columns);
            }
        }

        for (List<String> row : sheetTable) {
            ListIterator<String> colIterator = row.listIterator();

            while (colIterator.hasNext()) {
                String cellContent = colIterator.next();
                String formatString = "%-" + maxColumnLengths.get(colIterator.previousIndex()) + "s";
                String formattedCellContent = String.format(formatString, cellContent);

                fileContent.append(formattedCellContent);
                fileContent.append(" | ");
            }

            fileContent.append(System.lineSeparator());
        }

        /*
        for (int rowIndex = 0; rowIndex <= sheetTable.length - 1; rowIndex++) {
            for (int colIndex = 0; colIndex <= sheetTable[rowIndex].length - 1; colIndex++) {
                String formatString = "%-" + maxColumnLengths.get(colIndex) + "s";
                String cellContent = String.format(formatString, sheetTable[rowIndex][colIndex]);

                fileContent.append(cellContent);
                fileContent.append(" | ");
            }

            fileContent.append(System.lineSeparator());
        }

         */
    }

    private String getCellValueAndCharacteristics(Cell cell, Font font) {
        StringBuilder cellContent = new StringBuilder();

        switch (cell.getCellType()) {
            case NUMERIC:
                cellContent.append(cell.getNumericCellValue());
                break;
            case STRING:
                cellContent.append(cell.getStringCellValue());
                break;
            case FORMULA:
                cellContent.append(cell.getCellFormula());
                break;
            case BOOLEAN:
                cellContent.append(cell.getBooleanCellValue());
                break;
            case ERROR:
                cellContent.append(cell.getErrorCellValue());
                break;
            default:
                break;
        }

//        appendCellAllowedValues(cell, cellContent);
//        appendCellColor(cell, cellContent);
//        appendCellComment(cell, cellContent);
//        appendUnderline(font, cellContent);
//        appendBold(font, cellContent);

        return cellContent.toString();
    }

    private void appendCellAllowedValues(Cell cell, StringBuilder fileContent) {
        if (!cell.getSheet().getDataValidations().isEmpty()) {
            WorkbookEvaluatorProvider workbookEvaluatorProvider = (WorkbookEvaluatorProvider) cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            DataValidationEvaluator dataValidationEvaluator = new DataValidationEvaluator(cell.getSheet().getWorkbook(), workbookEvaluatorProvider);
            DataValidation cellDataValidation = dataValidationEvaluator.getValidationForCell(new CellReference(cell));
            if (cellDataValidation != null) {
//                String[] listValues = cellDataValidation.getValidationConstraint().getExplicitListValues();
                String[] listValues = cellDataValidation.getValidationConstraint().getExplicitListValues();
                fileContent.append("[");
                fileContent.append(String.join(", ", listValues));
                fileContent.append("] ");
            }
        }
    }

    private void appendUnderline(Font font, StringBuilder fileContent) {
        byte underline = font.getUnderline();
        if (underline == 1) {
            fileContent.append("Underline").append(" ");
        }
    }

    private void appendBold(Font font, StringBuilder fileContent) {
        if (font.getBold()) {
            fileContent.append("Bold").append(" ");
        }
    }

    private void appendCellColor(Cell cell, StringBuilder fileContent) {
        // (Foreground) Cell Color not set by Conditional Formatting
        XSSFColor foreColor = (XSSFColor) cell.getCellStyle().getFillForegroundColorColor();
        if (foreColor != null) {
            fileContent.append("#").append(foreColor.getARGBHex()).append(" ");
        }

        // (Background) Cell Color set by Conditional Formatting
        WorkbookEvaluatorProvider workbookEvaluatorProvider = (WorkbookEvaluatorProvider) cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
        ConditionalFormattingEvaluator conditionalFormattingEvaluator = new ConditionalFormattingEvaluator(cell.getSheet().getWorkbook(), workbookEvaluatorProvider);
        List<EvaluationConditionalFormatRule> matchingCFRules = conditionalFormattingEvaluator.getConditionalFormattingForCell(cell);
        for (EvaluationConditionalFormatRule evalCFRule : matchingCFRules) {
            ConditionalFormattingRule cFRule = evalCFRule.getRule();
            if (cFRule.getPatternFormatting() != null) {
                XSSFColor backColor = (XSSFColor) cFRule.getPatternFormatting().getFillBackgroundColorColor();
                fileContent.append("#").append(backColor.getARGBHex()).append(" ");
            } else if (cFRule.getColorScaleFormatting() != null) {
                XSSFColor[] colors = (XSSFColor[]) cFRule.getColorScaleFormatting().getColors();
                for (XSSFColor color : colors) {
                    fileContent.append("#").append(color.getARGBHex()).append(" ");
                }
            }
        }
    }

    private void appendCellComment(Cell cell, StringBuilder fileContent) {
        String separator = "Comment:";
        String flags = "Flags: ";
        if (cell.getCellComment() != null) {
            String cellComment = cell.getCellComment().getString().getString();
            int beginSubString = cellComment.indexOf(separator);
            String substring;
            if (cellComment.contains(separator)) {
                if (cellComment.contains(flags)) {
                    substring = cellComment.substring(beginSubString, cellComment.indexOf(flags)).replace("\n", "");
                } else {
                    substring = cellComment.substring(beginSubString).replace("\n", "");
                }
            } else {
                substring = cellComment.substring(0, cellComment.indexOf('\n'));
            }
            fileContent.append("{").append(substring).append("} ");
        }
    }

    private List<String> getAllExcelFiles() throws IOException {
        List<String> listOfExcelFiles;

        Path ectRootPath = Paths.get(searchDirectory);
        if (!ectRootPath.toFile().isDirectory()) {
            throw new IllegalArgumentException("Path must be a directory !");
        }

        try (Stream<Path> walk = Files.walk(ectRootPath)) {
            listOfExcelFiles = walk
                    .filter(p -> !p.toFile().isDirectory())
                    .map(Path::toString)
                    .filter(containsTarget.negate().and(isExcelFile))
                    .collect(Collectors.toList());
        }
        return listOfExcelFiles;
    }

    private String printFileEmptyOrDamagedMessage(String pathToExcel) {
        StringBuilder warning = new StringBuilder(Paths.get(pathToExcel).toFile().toString());

        if (Paths.get(pathToExcel).toFile().length() == 0) {
            warning.append(" can't be opened. This excel file may be empty !");
        } else {
            warning.append(" can't be opened. This excel file may be damaged !");
        }
        return warning.toString();
    }
}
