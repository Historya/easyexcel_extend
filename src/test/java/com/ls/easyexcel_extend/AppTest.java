package com.ls.easyexcel_extend;

import com.alibaba.excel.EasyExcel;
import com.ls.easyexcel_extend.comment.CommentModel;
import com.ls.easyexcel_extend.plugin.comment.ExcelHeadCommentHandler;
import com.ls.easyexcel_extend.plugin.select.ExcelSelectedHandler;
import com.ls.easyexcel_extend.selectd.SelectdModel;
import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;
import org.apache.commons.collections4.CollectionUtils;

/**
 * Unit test for simple App.
 */
public class AppTest 
    extends TestCase
{
    private final static String OUT_PATH = "D:\\my_work\\easyexcel_extend\\src\\test\\java\\com\\ls\\easyexcel_extend\\out\\";
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
    public AppTest( String testName )
    {
        super( testName );
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite()
    {
        return new TestSuite( AppTest.class );
    }

    /**
     * 测试批注
     */
    public void testComment(){
        EasyExcel.write()
                .head(CommentModel.class)
                .file(OUT_PATH+"测试EXCEL批准.xlsx")
                .sheet("sheet")
                .registerWriteHandler(new ExcelHeadCommentHandler<>(CommentModel.class))
                .doWrite(CollectionUtils.emptyCollection());
    }

    /**
     * 测试下拉选项
     */
    public void testSelect(){
        EasyExcel.write()
                .head(SelectdModel.class)
                .file(OUT_PATH+"测试EXCEL下拉.xlsx")
                .sheet("sheet")
                .registerWriteHandler(new ExcelSelectedHandler<>(SelectdModel.class))
                .doWrite(CollectionUtils.emptyCollection());
    }
}
