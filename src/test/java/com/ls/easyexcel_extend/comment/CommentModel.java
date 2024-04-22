package com.ls.easyexcel_extend.comment;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentStyle;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.ls.easyexcel_extend.base.Model;
import com.ls.easyexcel_extend.plugin.comment.ExcelComment;

/**
 * 批准示例excel模型
 * <br/>
 * date: 2024/4/21<br/>
 * version 0.1
 *
 * @author 10036<br />
 */
@ContentStyle(dataFormat = 49)
@HeadStyle(fillForegroundColor= 23)
public class CommentModel implements Model {

    @ExcelProperty(value = "标题1")
    @ExcelComment("标题1，xxxxxxxxxxxxxxxxxxxx")
    @ColumnWidth(value = 20)
    private String title1;

    @ExcelProperty(value = "标题2")
    @ExcelComment("标题2，xxxxxxxxxxxxxxxxxxxx")
    @ColumnWidth(value = 20)
    private String title2;
}
