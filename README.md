#快速进行excel导入导出的工具类

```
@Excel(
        name = "学生统计数据表",
        sheetName = "sheet1"
)
public class Student {

    private Long id;

    @CellValue(name = "名字",index = 0)
    private String name;

    @CellValue(name = "年龄",index = 1,suffix = "岁")
    private Integer age;

    @CellDate(name = "日期",index = 2,formatStr = "yyyy-MM-dd hh:mm:ss")
    private Date date;

    @CellBoolean(name = "是否删除",index = 3,tureValue = "是", falseValue = "否")
    private Boolean isDel;

    @CellDouble(name = "收入",index = 4, fixed = 2)
    private Double balance;

    @CellSelect(
            name = "选择",
            index = 5,
            keys = {"1","2","3"},
            values = {"选择1","选择2","选择3"}
    )
    private Integer select;

    @CellSelect(
            name = "选择2",
            index = 6,
            keys = {"a","b","c"},
            values = {"2选择1","2选择2","2选择3"}
    )
    private String select2;
```

注意属性field 基本数据类型 暂时只支持包装类型

@CellValue 通用注解
底层调用toString 方法 输出按照前缀和后缀拼接

属性 
    
    name 列名
    prefix 前缀
    suffix 后缀
    isMust 是否必须 如果对象为空会抛出异常 可以用异常捕获知道哪一行
    index cell的列号
    maxLen 最大长度
    
*这个注解的属性以下注解全都有下面的注解就省掉了

@CellBoolean 用于Boolean类型
    
    tureValue 为true的值
    falsValue 为false的值

@CellDate 用于Date
    
    formatStr 格式化
    
@CellDouble 用于Double、Float
    
    fixed 小数位后几位
    
@CellSelect 用于多选一的情况
    
    keys 键 对应对应下标values的值
    values 值 