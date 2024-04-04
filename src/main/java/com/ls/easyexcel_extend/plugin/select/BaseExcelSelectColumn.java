package com.ls.easyexcel_extend.plugin.select;

import lombok.Data;

/**
 * excel下拉选择列
 * @author ls
 * @version 1.0
 */
@Data
public abstract class BaseExcelSelectColumn {

    private ExcelSelected.Type type;

    Class<? extends ExcelDynamicDataSource> sourceHandel;


    String[] sourceParams;

    Integer parentColumnIndex;

    /**
     * 设置下拉框的起始行，默认为第二行
     */
    private int firstRow;

    /**
     * 设置下拉框的结束行，默认为最后一行
     */
    private int lastRow;
}