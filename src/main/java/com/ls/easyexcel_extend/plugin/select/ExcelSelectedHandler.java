package com.ls.easyexcel_extend.plugin.select;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.ls.easyexcel_extend.base.BaseHandler;
import com.ls.easyexcel_extend.base.Model;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

@Slf4j
public class ExcelSelectedHandler<E extends Model> implements BaseHandler<E>,SheetWriteHandler {
    private static final String PARAMETER_DEFINITIONS_SHEET_NAME = "系统参数";

    private final Map<Integer, ExcelSelectColumn> selectedResolveMap = new HashMap<>();

    private final Class<E> modelClass;

    public ExcelSelectedHandler(Class<E> modelClass) {
        super();
        this.modelClass = modelClass;
        this.init();
    }

    /**
     * Called before create the sheet
     */
    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (this.selectedResolveMap.isEmpty()) return;

        final Workbook workbook = writeWorkbookHolder.getWorkbook();
        final Sheet sheet = writeSheetHolder.getSheet();
        // 仅创建一个sheet用于存放下拉数据
        final AtomicReference<Sheet> definitionsSheet = new AtomicReference<>(ExcelSelectValidationUtil.createTmpSheet(workbook, PARAMETER_DEFINITIONS_SHEET_NAME));
        final AtomicInteger definitionsSheetStartColumn = new AtomicInteger(0);

        for (Map.Entry<Integer, ExcelSelectColumn> item : this.selectedResolveMap.entrySet()) {
            ExcelSelectColumn value = item.getValue();
            Integer index = item.getKey();

            if (value instanceof CascadeExcelSelectColumn) {
                CascadeExcelSelectColumn columnModel = (CascadeExcelSelectColumn) value;
                Map<String, String[]> source = columnModel.getSource();
                if(null == source || source.isEmpty()) continue;
                definitionsSheet.set(
                        ExcelSelectValidationUtil.addCascadeValidationToSheet(
                                workbook,
                                sheet,
                                definitionsSheet,
                                columnModel.getSource(),
                                definitionsSheetStartColumn,
                                columnModel.getParentColumnIndex(),
                                index,
                                columnModel.getFirstRow(),
                                columnModel.getLastRow()
                        )
                );
                continue;
            }

            if (value instanceof EasyExcelSelectColumn) {
                EasyExcelSelectColumn columnModel = (EasyExcelSelectColumn) value;
                String[] source = columnModel.getSource();
                if(null == source || source.length == 0) continue;
                definitionsSheet.set(
                        ExcelSelectValidationUtil.addSelectValidationToSheet(
                                workbook,
                                sheet,
                                definitionsSheet,
                                columnModel.getSource(),
                                definitionsSheetStartColumn,
                                index,
                                columnModel.getFirstRow(),
                                columnModel.getLastRow()
                        )
                );
            }
        }
    }

    /**
     * Called after the sheet is created
     */
    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (this.selectedResolveMap.isEmpty()) {
            return;
        }

        final Workbook workbook = writeWorkbookHolder.getWorkbook();
        final Sheet sheet = writeSheetHolder.getSheet();
        // 仅创建一个sheet用于存放下拉数据
        final AtomicReference<Sheet> definitionsSheet = new AtomicReference<>(ExcelSelectValidationUtil.createTmpSheet(workbook, PARAMETER_DEFINITIONS_SHEET_NAME));
        final AtomicInteger definitionsSheetStartColumn = new AtomicInteger(0);

        for (Map.Entry<Integer, ExcelSelectColumn> item : this.selectedResolveMap.entrySet()) {
            ExcelSelectColumn value = item.getValue();
            Integer index = item.getKey();

            if (value instanceof CascadeExcelSelectColumn) {
                CascadeExcelSelectColumn columnModel = (CascadeExcelSelectColumn) value;
                Map<String, String[]> source = columnModel.getSource();
                if(null == source || source.isEmpty()) continue;
                definitionsSheet.set(
                        ExcelSelectValidationUtil.addCascadeValidationToSheet(
                                workbook,
                                sheet,
                                definitionsSheet,
                                columnModel.getSource(),
                                definitionsSheetStartColumn,
                                columnModel.getParentColumnIndex(),
                                index,
                                columnModel.getFirstRow(),
                                columnModel.getLastRow()
                        )
                );
                continue;
            }

            if (value instanceof EasyExcelSelectColumn) {
                EasyExcelSelectColumn columnModel = (EasyExcelSelectColumn) value;
                String[] source = columnModel.getSource();
                if(null == source || source.length == 0) continue;
                definitionsSheet.set(
                        ExcelSelectValidationUtil.addSelectValidationToSheet(
                                workbook,
                                sheet,
                                definitionsSheet,
                                columnModel.getSource(),
                                definitionsSheetStartColumn,
                                index,
                                columnModel.getFirstRow(),
                                columnModel.getLastRow()
                        )
                );
            }
        }
    }

    private void init() {
        this.getAnalyzeModel();
    }

    /**
     * 分析 model
     */
    private void getAnalyzeModel() {
        Field[] fields = this.modelClass.getDeclaredFields();

        for (int i = 0, fieldsLength = fields.length; i < fieldsLength; i++) {
            Field field = fields[i];
            if (!field.isAnnotationPresent(ExcelSelected.class)) {
                continue;
            }

            ExcelSelected excelSelected = field.getAnnotation(ExcelSelected.class);
            int parentColumnIndex = excelSelected.parentColumnIndex();
            ExcelSelected.Type type = excelSelected.type();

            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            int colIndex = excelProperty.index();

            switch (type) {
                case CUSTOMER:
                    if (parentColumnIndex == -1) {
                        EasyExcelSelectColumn easyExcelSelectColumn = new EasyExcelSelectColumn();
                        easyExcelSelectColumn.setType(ExcelSelected.Type.CUSTOMER);
                        easyExcelSelectColumn.setLastRow(excelSelected.lastRow());
                        easyExcelSelectColumn.setFirstRow(excelSelected.firstRow());
                        easyExcelSelectColumn.setSourceHandel(excelSelected.sourceHandle());
                        easyExcelSelectColumn.setSourceParams(excelSelected.sourceParams());
                        this.selectedResolveMap.put((colIndex == -1 ? i : colIndex), easyExcelSelectColumn);
                    } else {
                        CascadeExcelSelectColumn cascadeExcelSelectColumn = new CascadeExcelSelectColumn();
                        cascadeExcelSelectColumn.setType(ExcelSelected.Type.CUSTOMER);
                        cascadeExcelSelectColumn.setLastRow(excelSelected.lastRow());
                        cascadeExcelSelectColumn.setFirstRow(excelSelected.firstRow());
                        cascadeExcelSelectColumn.setSourceHandel(excelSelected.sourceHandle());
                        cascadeExcelSelectColumn.setSourceParams(excelSelected.sourceParams());
                        cascadeExcelSelectColumn.setParentColumnIndex(excelSelected.parentColumnIndex());
                        this.selectedResolveMap.put((colIndex == -1 ? i : colIndex), cascadeExcelSelectColumn);
                    }
                    break;
                case SEQUENCE:
                    EasyExcelSelectColumn easyExcelSelectColumn = new EasyExcelSelectColumn();
                    easyExcelSelectColumn.setType(ExcelSelected.Type.SEQUENCE);
                    easyExcelSelectColumn.setLastRow(excelSelected.lastRow());
                    easyExcelSelectColumn.setFirstRow(excelSelected.firstRow());
                    easyExcelSelectColumn.setSource(excelSelected.source());
                    this.selectedResolveMap.put((colIndex == -1 ? i : colIndex), easyExcelSelectColumn);
                    break;
                default:
            }
        }

        this.analyzeSource();
    }

    private void analyzeSource() {
        //排序
        LinkedHashMap<Integer, ExcelSelectColumn> orderExcelSelectColumnMap = new LinkedHashMap<>();
        for (Map.Entry<Integer, ExcelSelectColumn> excelSelectColumnEntry : this.selectedResolveMap.entrySet()) {
            ExcelSelectColumn item = excelSelectColumnEntry.getValue();
            Integer key = excelSelectColumnEntry.getKey();
            if (item instanceof CascadeExcelSelectColumn && this.selectedResolveMap.containsKey(item.getParentColumnIndex())) {
                ExcelSelectColumn parentExcelSelectColumn = this.selectedResolveMap.get(item.getParentColumnIndex());
                if (!orderExcelSelectColumnMap.containsKey(item.getParentColumnIndex())) {
                    orderExcelSelectColumnMap.put(item.getParentColumnIndex(), parentExcelSelectColumn);
                }
            }
            orderExcelSelectColumnMap.put(key, item);
        }

        for (Map.Entry<Integer, ExcelSelectColumn> excelSelectColumnEntry : orderExcelSelectColumnMap.entrySet()) {
            ExcelSelectColumn item = excelSelectColumnEntry.getValue();

            if (ExcelSelected.Type.SEQUENCE.equals(item.getType())) continue;

            if (item instanceof EasyExcelSelectColumn) {
                this.easyExcelSelectAnalyzeSource((EasyExcelSelectColumn) item);
            } else if (item instanceof CascadeExcelSelectColumn) {
                ExcelSelectColumn parentExcelSelectColumn = this.selectedResolveMap.get(item.getParentColumnIndex());
                if (parentExcelSelectColumn instanceof CascadeExcelSelectColumn) {
                    CascadeExcelSelectColumn parentCascadeExcelSelectColumn = (CascadeExcelSelectColumn) parentExcelSelectColumn;
                    Map<String, String[]> parentSourceArr = parentCascadeExcelSelectColumn.getSource();
                    if (null == parentSourceArr || parentSourceArr.isEmpty()) {
                        continue;
                    }

                    final Set<String> dictionaryKeys = new HashSet<>();
                    parentSourceArr.forEach((index, values) -> Collections.addAll(dictionaryKeys, values));

                    this.cascadeExcelSelectAnalyzeSource((CascadeExcelSelectColumn) item, dictionaryKeys.toArray(new String[0]));
                    continue;
                }

                if (parentExcelSelectColumn instanceof EasyExcelSelectColumn) {
                    EasyExcelSelectColumn parentEasyExcelSelectColumn = (EasyExcelSelectColumn) parentExcelSelectColumn;
                    String[] parentSourceArr = parentEasyExcelSelectColumn.getSource();

                    if (null == parentSourceArr || parentSourceArr.length == 0) {
                        continue;
                    }

                    parentSourceArr = Arrays.stream(parentSourceArr).distinct().toArray(String[]::new);
                    this.cascadeExcelSelectAnalyzeSource((CascadeExcelSelectColumn) item, parentSourceArr);
                }
            }
        }
    }

    private void cascadeExcelSelectAnalyzeSource(CascadeExcelSelectColumn excelSelectColumn, String[] source) {
        try {
            ExcelDynamicDataSource excelDynamicDataSource = excelSelectColumn.getSourceHandel().newInstance();

            HashMap<String, String[]> sourceArr = new HashMap<>();

            for (String dictionaryKey : source) {
                String[] dictionaryKeySource = excelDynamicDataSource.getSource(new String[]{dictionaryKey});
                if (null == dictionaryKeySource || dictionaryKeySource.length == 0) {
                    continue;
                }
                sourceArr.put(dictionaryKey, dictionaryKeySource);
            }
            excelSelectColumn.setSource(sourceArr);
        } catch (Exception e) {
            log.error("解析EXCEL动态下拉框数据失败，失败信息：", e);
        }
    }

    private void easyExcelSelectAnalyzeSource(EasyExcelSelectColumn excelSelectColumn) {
        try {
            ExcelDynamicDataSource excelDynamicDataSource = excelSelectColumn.getSourceHandel().newInstance();
            String[] source = excelDynamicDataSource.getSource(excelSelectColumn.getSourceParams());
            excelSelectColumn.setSource(source);
        } catch (Exception e) {
            log.error("解析EXCEL动态下拉框数据失败，失败信息：", e);
        }
    }


    @Override
    public Class<E> getModelClass() {
        return this.modelClass;
    }

}