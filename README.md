# Excel 模板导出工具

一个无需 Spring 的 Java 8 工具库：读取 `.xlsx` 模板，将 JSON 数据填入占位符，并按数组数据复制模板行，生成新的 Excel 文件。

## 功能

- 支持 `${customer.name}` 这类对象路径和混合文本占位符。
- 支持 `${orders.productName}` 这类数组路径，按数组元素复制整行。
- 支持在同一个工作簿的多个工作表中重复使用数组。
- 保留单元格样式、行高、列宽、横向合并和相对公式。
- 缺失值或 `null` 写为空白；数字和布尔值在完整占位符中保留原始类型。
- 纯 Java 8 API，可用于命令行程序、Web 服务、定时任务等场景。

## 快速开始

### 1. 下载 Release 中的 JAR

在 [Releases](../../releases) 页面下载 `export-excel-1.0.0.jar`。这是一个供其他 Java 项目调用的库 JAR，没有 `main` 启动类，因此不能直接执行 `java -jar`。

推荐通过 Maven 引入：

```xml
<dependency>
    <groupId>com.example</groupId>
    <artifactId>export-excel</artifactId>
    <version>1.0.0</version>
</dependency>
```

如果只拿到了 Release JAR，也可以把它安装到本机 Maven 仓库：

```powershell
mvn install:install-file `
  "-Dfile=.\export-excel-1.0.0.jar" `
  "-DgroupId=com.example" `
  "-DartifactId=export-excel" `
  "-Dversion=1.0.0" `
  "-Dpackaging=jar"
```

工具依赖 Apache POI 和 Jackson。使用 Maven 时会自动解析传递依赖；手动使用 JAR 时，需要同时把这些依赖加入 classpath。

### 2. 准备模板和 JSON

模板单元格中使用 `${...}` 占位符。例如：

```text
A1: 客户：${customer.name}
A2: ${orders.productName}
B2: ${orders.quantity}
```

对应 JSON：

```json
{
  "customer": { "name": "张三" },
  "orders": [
    { "productName": "钢笔", "quantity": 2 },
    { "productName": "笔记本", "quantity": 5 }
  ]
}
```

`${orders.productName}` 和 `${orders.quantity}` 位于同一行时，导出器会按 `orders` 数组复制该模板行。

### 3. 调用导出接口

公开入口只有一个：

```java
public void export(InputStream template,
                   OutputStream output,
                   String json) throws IOException
```

完整示例：

```java
import com.example.excel.ExcelExporter;

import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class ExportDemo {
    public static void main(String[] args) throws Exception {
        Path templatePath = Paths.get("template.xlsx");
        Path resultPath = Paths.get("result.xlsx");
        String json = new String(
                Files.readAllBytes(Paths.get("data.json")),
                StandardCharsets.UTF_8);

        try (InputStream template = Files.newInputStream(templatePath);
             OutputStream output = Files.newOutputStream(resultPath)) {
            new ExcelExporter().export(template, output, json);
        }
    }
}
```

`template` 和 `output` 由调用方负责关闭。文件、上传文件流和 HTTP 响应流都可以适配这个接口。模板结构错误会抛出 `ExcelExportException`，JSON 或文件读写错误会抛出 `IOException`。

## 模板规则

| 模板内容 | 导出结果 |
| --- | --- |
| `${customer.name}` | 写入 `张三` |
| `客户：${customer.name}` | 写入 `客户：张三` |
| `${orders.productName}`、`${orders.quantity}` | 按 `orders` 数组复制整行 |
| `${payments.amount}` | 按 `payments` 数组复制整行 |
| `${tags}` | 对字符串数组逐行写入 |

注意事项：

- JSON 根节点必须是对象，路径使用点号分隔。
- 同一模板行只能引用一个数组；同一数组可以出现在多个工作表。
- 显式传入空数组时，对应模板行会被删除，下面的内容会上移。
- 数组字段不存在时会按普通空字段处理，模板行会保留。
- 完整占位符中的数字和布尔值保留对应 Excel 类型；混合文本始终写成字符串。
- 支持数组位于对象中，例如 `${sales.orders.product.name}`。
- 不支持嵌套列表、多行循环块、公式内部占位符和跨行合并的列表模板行。

## 运行示例

仓库自带模板和数据：

```text
examples/template.xlsx
examples/data.json
```

执行测试并打包：

```powershell
mvn clean test
mvn clean package
```

也可以在 IntelliJ IDEA 中运行 `src/test/java/com/example/excel/ExcelExportExample.java` 的 `main` 方法。示例输出写入 `target/example-output.xlsx`。

打包成功后，JAR 位于：

```text
target/export-excel-1.0.0.jar
```

## 手动 classpath（可选）

这是库 JAR，不能单独启动。若不使用 Maven，需要把 JAR 和 POI、Jackson 的全部运行时依赖一起放入 classpath。可以先复制 Maven 依赖：

```powershell
mvn dependency:copy-dependencies -DoutputDirectory=lib
```

Windows 示例：

```powershell
java -cp "export-excel-1.0.0.jar;lib/*;." ExportDemo
```

Linux/macOS 示例：

```bash
java -cp "export-excel-1.0.0.jar:lib/*:." ExportDemo
```

## 代码结构

- `ExcelExporter`：公开入口、JSON 解析和工作表遍历。
- `PlaceholderParser`：解析 `${...}` 占位符。
- `ValueResolver`：解析对象路径和数组路径。
- `ScalarRenderer`：替换普通单元格。
- `ListRenderer`：复制或删除列表模板行。
- `ExcelExporterTest`：导出结果和边界条件测试。

## 许可证

当前项目未附带单独许可证文件。如需在公开项目中长期使用，请根据分发方式补充许可证。
