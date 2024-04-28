package com.ls.easyexcel_extend.plugin.select;

import com.alibaba.excel.util.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Collectors;

/**
 * @author hp <br/>
 * <a href="https://www.bianchengbaodian.com/article/d539493521d9cf201294e9881b569168.html">来源</a>  <br/>
 */
public class ExcelSelectValidationUtil {

    private ExcelSelectValidationUtil() {
        throw new IllegalStateException("ExcelSelectValidationUtil class");
    }

    public static Sheet addCascadeValidationToSheet(
            Workbook workbook,
            Sheet sheet,
            AtomicReference<Sheet> tmpSheet,
            Map<String, String[]> options,
            AtomicInteger startCol,
            int parentCol,
            int selfCol,
            int startRow,
            int endRow
    ) {
        for (Map.Entry<String, String[]> entry : options.entrySet()) {
            String parentVal = formatNameManager(entry.getKey());
            String[] children = entry.getValue();
            if (null == children || children.length == 0) {
                continue;
            }
            int columnIndex = startCol.getAndIncrement();
            createDropdownElement(tmpSheet.get(), children, columnIndex);
            final String columnName = calculateColumnName(columnIndex + 1);
            final String formula = createFormulaForNameManger(tmpSheet.get(), children.length, columnName);
            createNameManager(workbook, parentVal, formula);
        }
        final String parentColumnName = calculateColumnName(parentCol + 1);
        final String indirectFormula = createIndirectFormula(parentColumnName, startRow + 1);
        createValidation(workbook, sheet, tmpSheet.get(), indirectFormula, selfCol, startRow, endRow);
        return tmpSheet.get();
    }


    public static Sheet createTmpSheet(Workbook workbook, String sheetName) {
        final String actualName = sheetName + workbook.getNumberOfSheets();
        final Sheet sheet = workbook.createSheet(actualName);
        /*
         * 关键:
         * 如果下拉一次写超过 windowSize, 内存数据被刷到磁盘, 导致内存中找不到此前写过的row对象
         * 见 new SXSSFSheet(windowSize);说明 setRandomAccessWindowSize(int)说明
         * -1 则不限, 请注意下拉数据量
         */
        if (sheet instanceof SXSSFSheet) {
            ((SXSSFSheet) sheet).setRandomAccessWindowSize(-1);
        }
        return sheet;
    }

    public static Sheet addSelectValidationToSheet(
            Workbook workbook,
            Sheet sheet,
            AtomicReference<Sheet> tmpSheet,
            String[] options,
            AtomicInteger startCol,
            int selfCol,
            int startRow,
            int endRow
    ) {
        final int columnIndex = startCol.getAndIncrement();
        String columnName = calculateColumnName(columnIndex + 1);
        final String formula = createFormulaForDropdown(tmpSheet.get(), options.length, columnName);
        createDropdownElement(tmpSheet.get(), options, columnIndex);
        createValidation(workbook, sheet, tmpSheet.get(), formula, selfCol, startRow, endRow);
        return tmpSheet.get();
    }


    private static void createDropdownElement(Sheet tmpSheet, String[] options, int columnIndex) {
        int rowIndex = 0;
        for (String val : options) {
            final int rIndex = rowIndex++;
            final Row row = Optional.ofNullable(tmpSheet.getRow(rIndex))
                    .orElseGet(() -> tmpSheet.createRow(rIndex));
            final Cell cell = row.createCell(columnIndex);
            cell.setCellValue(val);
        }
    }

    private static void createValidation(Workbook workbook, Sheet sheet, Sheet tmpSheet, String formula, int selfCol, int startRow, int endRow) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        final DataValidationConstraint constraint = helper.createFormulaListConstraint(formula);
        CellRangeAddressList addressList = new CellRangeAddressList(startRow, endRow, selfCol, selfCol);
        final DataValidation validation = helper.createValidation(constraint, addressList);
        validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        validation.setShowErrorBox(true);
        validation.setSuppressDropDownArrow(true);
        validation.createErrorBox("提示", "请输入下拉选项中的内容");
        sheet.addValidationData(validation);
        hideSheet(workbook, tmpSheet);
    }

    private static String createIndirectFormula(String columnName, int startRow) {
        final String format = "INDIRECT(CONCATENATE(\"_\",$%s%s))";
        return String.format(format, columnName, startRow);
    }

    private static String createFormulaForNameManger(Sheet tmpSheet, int size, String columnName) {
        final String format = "%s!$%s$%s:$%s$%s";
        return String.format(format, tmpSheet.getSheetName(), columnName, "1", columnName, size);
    }

    private static String createFormulaForDropdown(Sheet tmpSheet, int size, String columnName) {
        final String format = "=%s!$%s$%s:$%s$%s";
        return String.format(format, tmpSheet.getSheetName(), columnName, "1", columnName, size);
    }

    private static void createNameManager(Workbook workbook, String nameName, String formula) {
        //处理存在名称管理器复用的情况
        Name name = workbook.getName(nameName);
        if (name != null) {
            return;
        }
        name = workbook.createName();
        name.setNameName(nameName);
        name.setRefersToFormula(formula);
    }

    private static void hideSheet(Workbook workbook, Sheet sheet) {
        final int sheetIndex = workbook.getSheetIndex(sheet);
        if (sheetIndex > -1) {
            workbook.setSheetHidden(sheetIndex, true);
        }
    }

    private static String formatNameManager(String name) {
        if(StringUtils.isEmpty(name))  throw new IllegalArgumentException("错误的参数");
        name = name.trim();
        return "_" + name;
    }

    private static String calculateColumnName(int columnCount) {
        final int minimumExponent = minimumExponent(columnCount);
        final int base = 26;
        final int layers = (minimumExponent == 0 ? 1 : minimumExponent);
        final List<Character> sequence = new ArrayList<>();
        int remain = columnCount;
        for (int i = 0; i < layers; i++) {
            int step = (int) (remain / Math.pow(base, i) % base);
            step = step == 0 ? base : step;
            buildColumnNameSequence(sequence, step);
            remain = remain - step;
        }
        return sequence.stream()
                .map(Object::toString)
                .collect(Collectors.joining());
    }

    private static void buildColumnNameSequence(List<Character> sequence, int columnIndex) {
        final int capitalAAsIndex = 64;
        sequence.add(0, (char) (capitalAAsIndex + columnIndex));
    }

    private static int minimumExponent(int source) {
        final int base = 26;
        int exponent = 0;
        while (Math.pow(base, exponent) < source) {
            exponent++;
        }
        return exponent;
    }
}
