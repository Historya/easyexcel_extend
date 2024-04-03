package com.ls.easyexcel_extend.plugin.select;

import lombok.Data;

@Data
public abstract class ExcelSelectColumn {

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