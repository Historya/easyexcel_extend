package com.ls.easyexcel_extend.plugin.select;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.ls.easyexcel_extend.base.BaseHandler;
import com.ls.easyexcel_extend.base.Model;
import com.ls.easyexcel_extend.base.ModelField;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

/**
 * excel下拉选择处理
 *
 * @param <E> model
 * @author ls
 * @version 1.0
 * @see com.ls.easyexcel_extend.plugin.select.ExcelSelected 此处理类想要配合一起使用
 */
@Slf4j
public class ExcelSelectedHandler<E extends Model> implements BaseHandler<E>, SheetWriteHandler {
    private static final String PARAMETER_DEFINITIONS_SHEET_NAME = "系统参数";

    private final Map<Integer, BaseExcelSelectColumn> selectedResolveMap = new HashMap<>();

    private final Class<E> modelClass;

    public ExcelSelectedHandler(Class<E> modelClass) {
        super();
        this.modelClass = modelClass;
        this.init();
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

        for (Map.Entry<Integer, BaseExcelSelectColumn> item : this.selectedResolveMap.entrySet()) {
            BaseExcelSelectColumn value = item.getValue();
            Integer index = item.getKey();

            if (value instanceof CascadeExcelSelectColumn) {
                CascadeExcelSelectColumn columnModel = (CascadeExcelSelectColumn) value;
                Map<String, String[]> source = columnModel.getSource();
                if (null == source || source.isEmpty()) continue;
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

            if (value instanceof CommonExcelSelectColumn) {
                CommonExcelSelectColumn columnModel = (CommonExcelSelectColumn) value;
                String[] source = columnModel.getSource();
                if (null == source || source.length == 0) continue;
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
        this.analyzeModel();
    }

    /**
     * 分析 model
     */
    private void analyzeModel() {
        List<ModelField> fields = this.getModelFields();

        for (int i = 0, fieldsLength = fields.size(); i < fieldsLength; i++) {
            ModelField modelField = fields.get(i);
            Field field = modelField.getField();
            if (!field.isAnnotationPresent(ExcelSelected.class)) {
                continue;
            }

            ExcelSelected excelSelected = field.getAnnotation(ExcelSelected.class);
            int parentColumnIndex = excelSelected.parentColumnIndex();
            ExcelSelected.Type type = excelSelected.type();

            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            int colIndex = excelProperty.index();

            switch (type) {
                case SEQUENCE:
                    CommonExcelSelectColumn sequenceExcelSelectColumn = new CommonExcelSelectColumn();
                    sequenceExcelSelectColumn.setType(ExcelSelected.Type.SEQUENCE);
                    sequenceExcelSelectColumn.setLastRow(excelSelected.lastRow());
                    sequenceExcelSelectColumn.setFirstRow(excelSelected.firstRow());
                    sequenceExcelSelectColumn.setSource(excelSelected.source());

                    colIndex = colIndex == -1 ? modelField.getIndex() : colIndex;
                    this.selectedResolveMap.put(colIndex, sequenceExcelSelectColumn);
                    break;
                case CUSTOMER:
                    CommonExcelSelectColumn customerExcelSelectColumn = new CommonExcelSelectColumn();
                    customerExcelSelectColumn.setType(ExcelSelected.Type.CUSTOMER);
                    customerExcelSelectColumn.setLastRow(excelSelected.lastRow());
                    customerExcelSelectColumn.setFirstRow(excelSelected.firstRow());
                    customerExcelSelectColumn.setSourceHandel(excelSelected.sourceHandle());
                    customerExcelSelectColumn.setSourceParams(excelSelected.sourceParams());
                    this.selectedResolveMap.put((colIndex == -1 ? i : colIndex), customerExcelSelectColumn);
                    break;
                case CASCADE:
                    if (parentColumnIndex != -1) {
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
                default:
            }
        }

        this.computeSourceData();
    }

    /**
     * 计算下拉资源
     */
    private void computeSourceData() {
        //排序
        LinkedHashMap<Integer, BaseExcelSelectColumn> orderExcelSelectColumnMap = this.computeInitSourceDataOrder();

        this.initSourceData(orderExcelSelectColumnMap);
    }

    /**
     * 初始化下拉资源
     *
     * @param orderExcelSelectColumnMap 有序的下拉选择列信息
     */
    private void initSourceData(LinkedHashMap<Integer, BaseExcelSelectColumn> orderExcelSelectColumnMap) {
        for (Map.Entry<Integer, BaseExcelSelectColumn> excelSelectColumnEntry : orderExcelSelectColumnMap.entrySet()) {
            BaseExcelSelectColumn item = excelSelectColumnEntry.getValue();

            if (ExcelSelected.Type.SEQUENCE.equals(item.getType())) continue;

            if (item instanceof CommonExcelSelectColumn) {
                this.getEasyExcelSelectSourceData((CommonExcelSelectColumn) item);
                continue;
            }

            if (item instanceof CascadeExcelSelectColumn) {
                BaseExcelSelectColumn parentBaseExcelSelectColumn = this.selectedResolveMap.get(item.getParentColumnIndex());
                if (parentBaseExcelSelectColumn instanceof CascadeExcelSelectColumn) {
                    CascadeExcelSelectColumn parentCascadeExcelSelectColumn = (CascadeExcelSelectColumn) parentBaseExcelSelectColumn;
                    Map<String, String[]> parentSourceArr = parentCascadeExcelSelectColumn.getSource();
                    if (null == parentSourceArr || parentSourceArr.isEmpty()) {
                        continue;
                    }

                    final Set<String> dictionaryKeys = new HashSet<>();
                    parentSourceArr.forEach((index, values) -> Collections.addAll(dictionaryKeys, values));

                    this.getCascadeExcelSelectSourceData((CascadeExcelSelectColumn) item, dictionaryKeys.toArray(new String[0]));
                    continue;
                }

                if (parentBaseExcelSelectColumn instanceof CommonExcelSelectColumn) {
                    CommonExcelSelectColumn parentEasyExcelSelectColumn = (CommonExcelSelectColumn) parentBaseExcelSelectColumn;
                    String[] parentSourceArr = parentEasyExcelSelectColumn.getSource();

                    if (null == parentSourceArr || parentSourceArr.length == 0) continue;

                    parentSourceArr = Arrays.stream(parentSourceArr).distinct().toArray(String[]::new);
                    this.getCascadeExcelSelectSourceData((CascadeExcelSelectColumn) item, parentSourceArr);
                }
            }
        }
    }

    /**
     * 计算初始化顺序
     *
     * @return 计算顺序后的下拉选择列信息列表
     */
    private LinkedHashMap<Integer, BaseExcelSelectColumn> computeInitSourceDataOrder() {
        LinkedHashMap<Integer, BaseExcelSelectColumn> orderExcelSelectColumnMap = new LinkedHashMap<>();
        for (Map.Entry<Integer, BaseExcelSelectColumn> excelSelectColumnEntry : this.selectedResolveMap.entrySet()) {
            BaseExcelSelectColumn item = excelSelectColumnEntry.getValue();
            Integer key = excelSelectColumnEntry.getKey();
            if (item instanceof CascadeExcelSelectColumn && this.selectedResolveMap.containsKey(item.getParentColumnIndex())) {
                BaseExcelSelectColumn parentBaseExcelSelectColumn = this.selectedResolveMap.get(item.getParentColumnIndex());
                if (!orderExcelSelectColumnMap.containsKey(item.getParentColumnIndex())) {
                    orderExcelSelectColumnMap.put(item.getParentColumnIndex(), parentBaseExcelSelectColumn);
                }
            }
            orderExcelSelectColumnMap.put(key, item);
        }
        return orderExcelSelectColumnMap;
    }

    private void getCascadeExcelSelectSourceData(CascadeExcelSelectColumn excelSelectColumn, String[] source) {
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

    private void getEasyExcelSelectSourceData(CommonExcelSelectColumn excelSelectColumn) {
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