# Export-Excel-With-Image
使用Java POI编写的导出图片到Excel的工具

图片宽度固定，高度会自适应

```java
  sheet.setColumnWidth(i, (30 * 256));
```
这里的30\*256是列宽度 代表30个字符数 建议设置好了打一张空表格看一看是多少像素（这里的字符数一般是**宋体**字号为**11**的情况下一行最多显示**英文状态**下的数字的个数）

```java
rowHeight = (short) ((245D / 96D) * 72 * 20 * ((double) height / (double) width));
```
这里的行高是根据列宽计算的，其中245就是30字符的列宽对应的245像素宽度，再根据图片宽高比例和换算公式计算出需要设置的行高磅数

## Maven
```xml
    <dependency>
	<groupId>org.apache.poi</groupId>
	<artifactId>poi</artifactId>
	<version>3.10-FINAL</version>
    </dependency>
    <dependency>
	<groupId>org.apache.poi</groupId>
	<artifactId>poi-ooxml</artifactId>
	<version>3.10-FINAL</version>
    </dependency>
    <dependency>
	<groupId>org.apache.poi</groupId>
	<artifactId>poi-ooxml-schemas</artifactId>
	<version>3.10-FINAL</version>
    </dependency>   
```
