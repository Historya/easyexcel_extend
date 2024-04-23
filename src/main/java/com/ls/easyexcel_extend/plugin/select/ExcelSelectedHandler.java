package com.ls.easyexcel_extend.plugin.select;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.ls.easyexcel_extend.base.BaseHandler;
import com.ls.easyexcel_extend.base.Model;
import com.ls.easyexcel_extend.base.ModelField;
import com.ls.easyexcel_extend.exception.ExcelHandlerException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

/**
 * excel下拉选择处理
 *
 * @param <E> model
 * @author ls
 * @version 1.0
 * @see com.ls.easyexcel_extend.plugin.select.ExcelSelected 此处理类需要配合一起使用
 */
@Slf4j
public class ExcelSelectedHandler<E extends Model> implements BaseHandler<E>, SheetWriteHandler {
    private static final String PARAMETER_DEFINITIONS_SHEET_NAME = "系统参数";

    private final Map<Integer,BaseExcelSelectColumn> selectedResolveMap = new ConcurrentHashMap<>();

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
        this.computeSourceData();
    }

    /**
     * 分析 model
     */
    private void analyzeModel() {
        List<ModelField> fields = this.getModelFields();

        if (fields.isEmpty()) return;

        for (int i = 0, fieldsLength = fields.size(); i < fieldsLength; i++) {
            ModelField modelField = fields.get(i);
            Field field = modelField.getField();

            //没有添加@ExcelSelected注解忽略
            if (!field.isAnnotationPresent(ExcelSelected.class)) continue;

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
                        cascadeExcelSelectColumn.setType(ExcelSelected.Type.CASCADE);
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
    }

    /**
     * 计算下拉资源
     */
    private void computeSourceData() {
        if (this.selectedResolveMap.isEmpty()) return;

        this.selectedResolveMap.values().parallelStream().forEach(e -> {
            if (!e.isSourceEmpty()) return;
            doComputeSourceData(e);
        });
    }

    /**
     * 计算下拉资源
     * @param excelSelectColumn 计算列
     */
    private void doComputeSourceData(BaseExcelSelectColumn excelSelectColumn) {
        switch (excelSelectColumn.getType()) {
            case CUSTOMER:
                this.computeCustomerColumn((CommonExcelSelectColumn) excelSelectColumn);
                break;
            case CASCADE:
                BaseExcelSelectColumn parentExcelSelectColumn = this.selectedResolveMap.getOrDefault(excelSelectColumn.getParentColumnIndex(),null);
                if(null == parentExcelSelectColumn) return;

                if (parentExcelSelectColumn.isSourceEmpty()) {
                    this.doComputeSourceData(parentExcelSelectColumn);
                }

                String[] parentSourceArr = null;

                if (parentExcelSelectColumn instanceof CascadeExcelSelectColumn) {
                    Map<String, String[]> parentSource = ((CascadeExcelSelectColumn) parentExcelSelectColumn).getSource();

                    if(null ==  parentSource || parentSource.isEmpty()) break;

                    Set<String> dictionaryKeys = new HashSet<>();
                    parentSource.forEach((index, values) -> Collections.addAll(dictionaryKeys, values));
                    parentSourceArr = dictionaryKeys.stream().distinct().toArray(String[]::new);

                } else {
                    parentSourceArr = ((CommonExcelSelectColumn) parentExcelSelectColumn).getSource();
                }

                if(null ==  parentSourceArr || parentSourceArr.length == 0) break;

                this.computeCascadeColumn((CascadeExcelSelectColumn)excelSelectColumn,parentSourceArr);
                break;
            case SEQUENCE:
            default:
        }
    }

    /**
     * 计算联级列下拉资源
     * @param excelSelectColumn 计算列
     * @param parentSourceArr 参数
     */
    private void computeCascadeColumn(CascadeExcelSelectColumn excelSelectColumn, String[] parentSourceArr) {

        try {
            ExcelDynamicDataSource excelDynamicDataSource = excelSelectColumn.getSourceHandel().newInstance();

            HashMap<String, String[]> sourceArr = new HashMap<>();

            for (String dictionaryKey : parentSourceArr) {
                String[] dictionaryKeySource = excelDynamicDataSource.getSource(new String[]{dictionaryKey});
                sourceArr.put(dictionaryKey, dictionaryKeySource);
            }

            excelSelectColumn.setSource(sourceArr);
        } catch (Exception e) {
            throw new ExcelHandlerException("解析EXCEL动态下拉框数据失败", e);
        }

    }

    /**
     *计算自定义列下拉资源
     * @param excelSelectColumn 计算列
     */
    private void computeCustomerColumn(CommonExcelSelectColumn excelSelectColumn) {
        try {
            ExcelDynamicDataSource excelDynamicDataSource = excelSelectColumn.getSourceHandel().newInstance();
            String[] source = excelDynamicDataSource.getSource(excelSelectColumn.getSourceParams());
            excelSelectColumn.setSource(source);
        } catch (Exception e) {
            throw new ExcelHandlerException("解析EXCEL动态下拉框数据失败", e);
        }
    }

    @Override
    public Class<E> getModelClass() {
        return this.modelClass;
    }

}