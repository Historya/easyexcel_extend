package com.ls.easyexcel_extend.plugin.select;

/**
 * 动态生成的下拉框可选数据
 * @author ls
 * @version 1.0
 */
public interface ExcelDynamicDataSource {
    /**
     * 获取动态生成的下拉框可选数据
     * @return 动态生成的下拉框可选数据
     */
    String[] getSource(Object[] param);
}