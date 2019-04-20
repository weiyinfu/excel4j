入口类ExcelUtil

```java
ExcelUtil.exportExcel(new String[]{"姓名", "年龄"},
        new String[]{"name", "age"},
        getList()).write(new File("haha.xls"));
List<User> a = ExcelUtil.importExcel(
        new FileInputStream("haha.xls"),
        new String[]{"姓名", "年龄"},
        new String[]{"name", "age"},
        User.class);
```