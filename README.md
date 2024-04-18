# easyexcel 扩展插件
--
基于 easyExcel **3.3.x** 开发

## 1.注解批注功能
##### 代码：
``` java
public class StationExcelDTO {
  
      @ExcelProperty(value = "加油站",index = 0)
      @ExcelCommentAnnotation("必填")
      @ColumnWidth(value = 20)
      @ContentStyle(dataFormat = 49)
      private String name;
      
      @ExcelProperty(value = "所属机构",index = 1)
      @ExcelCommentAnnotation("必填,填入系统管理>部门管理菜单下的部门名称")
      @ColumnWidth(value = 30)
      @ContentStyle(dataFormat = 49)
      private String organizeId;
  
      //.......
}
```
   
```java
@GetMapping("/template")
public void templateExcel(HttpServletResponse response, HttpServletRequest request) throws IOException {
    // 文件名字
    String fileName = "加油站导入模板";
    CommonUtil.exportFileResponseHeader(response, request, fileName, "application/vnd.ms-excel");
    try (ServletOutputStream os = response.getOutputStream()) {
        EasyExcel.write(os, StationExcelDTO.class)
                .sheet("加油站信息")
                .registerWriteHandler(new ExcelHeadCommentHandler<>(StationExcelDTO.class))
                .doWrite(null);
    }
}
```
##### 效果：
![批注效果](doc/imgs/comment.png)

## 2.下拉选项
###### 说明
```java
public @interface ExcelSelected {
    /**
     * 类型
     */
    Type type();

    /**
     * 固定下拉内容
     */
    String[] source() default {};

    /**
     * 动态数据类
     * 动态下拉内容
     */
    Class<? extends ExcelDynamicDataSource> sourceHandle();

    /**
     * 动态数据 参数
     */
    String[] sourceParams() default {};

    /**
     *父索引
     */
    int parentColumnIndex() default -1;

    /**
     * 设置下拉框的起始行，默认为第二行
     */
    int firstRow() default 1;

    /**
     * 设置下拉框的结束行，默认为最后一行
     * 65536
     */
    int lastRow() default 65536;

    static enum Type {
        SEQUENCE,
        CUSTOMER
    }
}
```

```java
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

```

###### 代码

```java
public class OliInApplyExcelDTO {

    @ColumnWidth(value = 20)
    @ContentStyle(dataFormat = 49)
    @ExcelCommentAnnotation(value = "单号不能为空,字符长度14")
    @ExcelProperty(value = "单号", index = 0)
    private String code;

    @ColumnWidth(value = 20)
    @ContentStyle(dataFormat = 49)
    @ExcelCommentAnnotation("进油日期,yyyy-MM-dd")
    @ExcelProperty(value = "进油日期", index = 1)
    @DateTimeFormat(value = DateUtil.SIMPLE_DAY_DATE_FORMAT)
    private Date dateTime;

    @ExcelSelected(type = ExcelSelected.Type.CUSTOMER, sourceClass = StationDataSource.class, firstRow = 2)
    @ExcelCommentAnnotation("油站，不能为空,字符长度1-128")
    @ColumnWidth(value = 20)
    @ContentStyle(dataFormat = 49)
    @ExcelProperty(value = "油站", index = 2)
    private String stationName;

    /**
     * 备注
     */
    @ExcelCommentAnnotation("备注,字符长度1~1000字")
    @ExcelProperty(value = "备注", index = 3)
    @ColumnWidth(value = 20)
    @ContentStyle(dataFormat = 49)
    private String remark;

    @ExcelCommentAnnotation("商品类型,字符长度1~120字")
    @ExcelProperty(value = {"进油明细", "商品类型"}, index = 4)
    @ColumnWidth(value = 20)
    @ContentStyle(dataFormat = 49)
    @ExcelSelected(type = ExcelSelected.Type.CUSTOMER, sourceClass = ItemClassDataSource.class, firstRow = 2)
    private String itemClassName;

    @ExcelCommentAnnotation("商品名称,字符长度1-128")
    @ExcelSelected(type = ExcelSelected.Type.CUSTOMER, sourceClass = ItemDataSource.class, parentColumnIndex = 4, firstRow = 2)
    @ColumnWidth(value = 20)
    @ContentStyle(dataFormat = 49)
    @ExcelProperty(value = {"进油明细", "商品名称"}, index = 5)
    private String itemName;

    /**
     * 商品数量
     */
    @ExcelCommentAnnotation("进油数量,不能为空")
    @ExcelProperty(value = {"进油明细", "进油数量"}, index = 6)
    @ColumnWidth(value = 15)
    @ContentStyle(dataFormat = 49)
    private String itemQuantity;

    //................
}
```

```java
public class ItemClassDataSource implements ExcelDynamicDataSource {

    /**
     * 获取动态生成的下拉框可选数据
     *
     * @param params 参数
     * @return 动态生成的下拉框可选数据
     */
    @Override
    public String[] getSource(String[] params) {
        ItemClassService itemClassService = SpringUtil.getBean(ItemClassService.class);
        if(null == itemClassService) return new String[0];

        List<EntityItemClassAO> itemClassList = itemClassService.getItemClassListByPid(StaticDefine.ROOT_NODE).getData();
        if(CollectionUtils.isEmpty(itemClassList)) return new String[0];

        return itemClassList.stream().map(EntityItemClassAO::getName).distinct().toArray(String[]::new);
    }
}
```

```java
@GetMapping("/template")
public void templateExcel(HttpServletResponse response, HttpServletRequest request) throws IOException {
    // 文件名字
    String fileName = "加油站导入模板";
    CommonUtil.exportFileResponseHeader(response, request, fileName, "application/vnd.ms-excel");
    try (ServletOutputStream os = response.getOutputStream()) {
        EasyExcel.write(os, StationExcelDTO.class)
                .sheet("加油站信息")
                .registerWriteHandler(new ExcelSelectedHandler<>(StationExcelDTO.class))
                .doWrite(null);
    }
}
```

###### 效果：
![下拉参数表](doc/imgs/select_options.png)

![下拉效果](doc/imgs/select_options_show.png)