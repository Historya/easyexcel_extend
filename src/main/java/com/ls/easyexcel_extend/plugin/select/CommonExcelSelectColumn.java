package com.ls.easyexcel_extend.plugin.select;

import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * 常规下拉选择列
 * @author ls
 * @version 1.0
 */
@EqualsAndHashCode(callSuper = true)
@Data
public final class CommonExcelSelectColumn extends BaseExcelSelectColumn {
    /**
     * 下拉内容
     */
    private String[] source;
}