package com.ls.easyexcel_extend.plugin.select;

/**
 * 动态生成的下拉框可选数据
 * @author ls
 * @version 1.0
 * @see <a href="https://www.bianchengbaodian.com/article/d539493521d9cf201294e9881b569168.html">参考</a><br/>
 */
public interface ExcelDynamicDataSource {
    /**
     * 获取动态生成的下拉框可选数据
     * @return 动态生成的下拉框可选数据
     */
    String[] getSource(String[] param);
}