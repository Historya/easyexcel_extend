package com.ls.easyexcel_extend.selectd;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentStyle;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.ls.easyexcel_extend.base.Model;
import com.ls.easyexcel_extend.plugin.select.ExcelSelected;

/**
 * 示例excel下拉选项模型
 * <br/>
 * date: 2024/4/21<br/>
 * version 0.1
 *
 * @author ls<br />
 */
@ContentStyle(dataFormat = 49)
@HeadStyle(fillForegroundColor= 23)
public class SelectdModel implements Model {

    @ExcelProperty(value = "标题1",index = 0)
    @ExcelSelected(type = ExcelSelected.Type.SEQUENCE, source = {"一", "二", "三"})
    @ColumnWidth(value = 20)
    private String title1;

    @ExcelSelected(type = ExcelSelected.Type.CASCADE, sourceHandle = AssociationDataSource.class,parentColumnIndex = 0)
    @ExcelProperty(value = "标题2")
    @ColumnWidth(value = 20)
    private String title2;

    @ExcelSelected(type = ExcelSelected.Type.CUSTOMER, sourceHandle = ParamsDataSource.class,sourceParams = {"1","2","3"})
    @ExcelProperty(value = "标题3")
    @ColumnWidth(value = 20)
    private String title3;
}
