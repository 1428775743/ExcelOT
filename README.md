#快速进行excel导入导出的工具类

```
@ExcelBean(
    name = "导出名称",
    Excel.v2003,
    sheetName = "自定义名称(页)"
)
class bean{
    
    @Cell(
        name = "列名"
    )
    String field;
}
```

有三个使用用例在目录下， 可以非常方便的使用