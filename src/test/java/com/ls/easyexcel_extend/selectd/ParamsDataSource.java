package com.ls.easyexcel_extend.selectd;

import com.ls.easyexcel_extend.plugin.select.ExcelDynamicDataSource;

/**
 * <br/>
 * date: 2024/4/21<br/>
 * version 0.1
 *
 * @author 10036<br />
 */
public class ParamsDataSource implements ExcelDynamicDataSource {
    @Override
    public String[] getSource(String[] param) {
        return param;
    }
}
