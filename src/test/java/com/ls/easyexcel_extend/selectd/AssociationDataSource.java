package com.ls.easyexcel_extend.selectd;

import com.ls.easyexcel_extend.plugin.select.ExcelDynamicDataSource;

public class AssociationDataSource implements ExcelDynamicDataSource {
    @Override
    public String[] getSource(String[] param) {
        String key = param[0];
        if("一".equals(key)){
            return new String[]{"1-1","1-2"};
        }
        if("二".equals(key)){
            return new String[]{"2-1","2-2"};
        }
        if("三".equals(key)){
            return new String[]{"3-1","3-2"};
        }
        return null;
    }
}