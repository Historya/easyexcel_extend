package com.ls.easyexcel_extend.plugin.select;

import com.ls.easyexcel_extend.exception.ExcelHandlerException;
import lombok.Data;

/**
 * excel下拉选择列
 * @author ls
 * @version 1.0
 */
@Data
public abstract class BaseExcelSelectColumn<T> {

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


    /**
     * 下拉内容
     */
    private T source;

    /**
     * 判断source是否是空的
     * @return 是否
     */
    public boolean isSourceEmpty(){
        throw new ExcelHandlerException("please override this method");
    }
}