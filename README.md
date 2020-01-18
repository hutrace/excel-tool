
# Excel工具包说明

> 读写Excel的工具包
> 依赖apache的poi，封装读写方法，让读写Excel方便
> 所有方法均已封装至org.hutrace.exceltool.Excel工具类，可直接快速调用。

__简单使用示例__
* Read:
```java
	List<Map<String, Object>> data = Excel.readToMap("file path");
```

* Write:
```java
	Excel.writeToFile(data, "file path");
```

__org.hutrace.exceltool.Excel类封装了全面的使用方法，如果你没有找到想要的使用方法，你可以直接使用Reader和Writer__

* Read:
```java
	Reader reader = new Reader();
	List<Map<String, Object>> data = reader.toMap("file path");
```

* Write:
```java
	Writer writer = new Writer();
	writer.toFile(data, "file path");
```

__如果你在Reader/Writer中没有找到你需要使用的方法，你可以自己进行扩展，扩展非常简单__
* 第一步:

> 继承Reader/Writer

* 第二步:

> 增加你需要的方法,例如在Reader中新增toMap(byte[] bytes)方法
```java
    class CustomReader extends Reader{
        public List<Map<String, Object>> toMap(byte[] bytes, ExcelType type) throws IOException, NotFoundSheetException {
            // 你也可以自己定义解析器，返回不同类型的数据。
            MapResolver resolver = new MapResolver();
            reading(new ByteArrayInputStream(bytes), type, resolver);
            return resolver.data();
        }
    }
```

* 第三步:

> 如果你需要不同的数据解析器（当前已有Map解析器、Map别名解析器、JavaBean解析器），你可以选择继承AbstractReaderResolver抽象类，也可以选择实现ReaderResolver接口，AbstractReaderResolver内部实现了一些公共的方法，你可以选择使用它。

