package com.ls.easyexcel_extend.plugin.comment;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.ls.easyexcel_extend.base.BaseHandler;
import com.ls.easyexcel_extend.base.Model;
import com.ls.easyexcel_extend.base.ModelField;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.lang.reflect.Field;
import java.util.List;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 表头批注处理
 *
 * @param <E> 定义了表头信息的类
 * @see com.ls.easyexcel_extend.plugin.comment.ExcelComment 此处理类想要配合一起使用
 * @author ls
 * @version 1.0
 */
@Slf4j
public class ExcelHeadCommentHandler<E extends Model> implements BaseHandler<E>, CellWriteHandler {

    private final Class<E> modelClass;

    /**
     * 缓存表头批注信息
     */
    private final ConcurrentHashMap<Integer, ExcelComment> excelHeadCommentMap = new ConcurrentHashMap<>();

    public ExcelHeadCommentHandler(Class<E> modeClass) {
        this.modelClass = modeClass;
        this.getNotationMap();
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
                                 List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (Boolean.FALSE.equals(isHead) || this.excelHeadCommentMap.isEmpty() ||
                !this.excelHeadCommentMap.containsKey(cell.getColumnIndex())) return;
        List<String> headNameList = head.getHeadNameList();

        //只在表头最后一行添加
        if(headNameList.size() > 1 && relativeRowIndex != headNameList.size() -1 ) return;

        // 批注内容
        ExcelComment excelComment = this.excelHeadCommentMap.get(cell.getColumnIndex());
        // 创建绘图对象
        XSSFClientAnchor anchor = new XSSFClientAnchor();
        int rowIndex = cell.getRowIndex();
        //计算合并单元开始位置
        if(headNameList.size() > 1){
            String cellKey = headNameList.get(cell.getRowIndex());
            for (int i = cell.getRowIndex() -1; i >=0; i--) {
                if(cellKey.equals(headNameList.get(i))){
                    rowIndex--;
                }else {
                    break;
                }
            }
        }

        int columnIndex = cell.getColumnIndex();
        anchor.setRow1(rowIndex);
        anchor.setCol1(columnIndex);
        anchor.setRow2(rowIndex + excelComment.getRemarkColumnWide());
        anchor.setCol2(columnIndex + excelComment.getRemarkRowHigh());
        Sheet sheet = writeSheetHolder.getSheet();
        Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch();
        Comment comment = drawingPatriarch.createCellComment(anchor);
        comment.setString(new XSSFRichTextString(excelComment.getRemarkValue()));
        cell.setCellComment(comment);
    }

    /**
     * 获取批注信息
     */
    private void getNotationMap() {
        List<ModelField> modelFieldList = this.getModelFields();
        for (ModelField column : modelFieldList) {
            Field field = column.getField();
            ExcelComment excelComment = new ExcelComment();
            com.ls.easyexcel_extend.plugin.comment.ExcelComment excelCommentAnnotation = field.getAnnotation(com.ls.easyexcel_extend.plugin.comment.ExcelComment.class);
            if(null == excelCommentAnnotation){
                continue;
            }
            excelComment.setRemarkValue(excelCommentAnnotation.value());
            excelComment.setRemarkColumnWide(excelCommentAnnotation.remarkColumnWide());
            excelComment.setRemarkRowHigh(excelCommentAnnotation.remarkRowHigh());
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            int colIndex = excelProperty.index();
            excelComment.setColumn(excelProperty.index());
            this.excelHeadCommentMap.put((colIndex == -1 ? column.getIndex() : colIndex), excelComment);
        }
    }

    @Override
    public Class<E> getModelClass() {
        return this.modelClass;
    }

    @Data
    static class ExcelComment {

        /**
         * 列号
         */
        private Integer column;

        /**
         * 批注值
         */
        private String remarkValue;

        /**
         * 批注行高
         */
        int remarkRowHigh;

        /**
         * 批注列宽
         */
        int remarkColumnWide;

        /**
         * 批注所在行
         */
        int row;
    }
}