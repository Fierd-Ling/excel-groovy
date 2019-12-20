
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;


/**
 * excel 解析成对应的groovy脚本
 *
 * @author lingyun.zhong@hand-china.com
 * @date 2019/12/9 18:11
 **/
public class ExcelGroovy {

    public static void main(String[] args) throws IOException {
        createFile("lingyun.zhong@hand-china.com",
                "C:\\Users\\TF\\Desktop\\一步制造\\999 一步云制造产品研发\\35 表设计\\一步云制造 产品研发 LMDS服务表设计 V1.2.xls",
                "资源TPM设置");
    }


    private static void createFile(String email, String fileUrl, String... sheetName) throws IOException {
        String[] sheetNames = sheetName.clone();
        InputStream inputStream = new FileInputStream(fileUrl);
        Workbook workbook = new HSSFWorkbook(inputStream);
        List<Sheet> sheets = new ArrayList<>();
        for (String name : sheetNames) {
            sheets.add(workbook.getSheet(name));
        }
        for (Sheet sheet : sheets) {
            getCellValue(sheet, email);
        }
    }

    private static void getCellValue(Sheet sheet, String email) throws IOException {
        Map<String, Map<String, String>> cellMap = new HashMap<>(16);
        Map<String, List<String>> listMap = new HashMap<>(16);
        String primaryKey = null;
        String tableName = null;
        String tableDes = null;
        int rowCount = sheet.getLastRowNum();
        // excel 实际开始的行是第二行开始
        for (int r = 1; r < rowCount; r++) {
            // excel 的列截止的是第七列，开始的是第二列，特殊情况比如索引会超过七列
            Row row = sheet.getRow(r);
            if (row == null) {
                // 如果row等于空，就是字段和索引部分已经便利完毕
                break;
            }
            int cellCount = 7;
            Map<String, String> cellValueMap = new HashMap<>(8);
            String mapKey = null;
            for (int c = 1; c < cellCount; c++) {
                Cell cell = row.getCell(c);
                String cellValue;
                try {
                    cellValue = cell.getStringCellValue();
                } catch (IllegalStateException e) {
                    // 数字会直接导致获取失败，只要小数点前面的
                    cellValue = String.valueOf(cell.getNumericCellValue()).split("[.]")[0];
                }
                if (c == 1) {
                    mapKey = cellValue;
                    if ("".equals(cellValue)) {
                        break;
                    }
                    if ("开发简要设计".equals(cellValue)) {
                        break;
                    }
                    if ("数据量估算".equals(cellValue)) {
                        break;
                    }
                    if ("WHO字段".equals(cellValue)) {
                        break;
                    }
                    if ("字段名".equals(cellValue)) {
                        break;
                    }
                    if ("索引类型".equals(cellValue)) {
                        break;
                    }
                    // 索引处理
                    if ("主键".equals(mapKey)) {
                        primaryKey = row.getCell(2).getStringCellValue();
                        break;
                    }
                    if ("唯一性索引".equals(mapKey) || "普通索引".equals(mapKey)) {
                        index(mapKey, listMap, row);
                        break;
                    }
                    // 表头处理
                    if ("表名/描述".equals(mapKey)) {
                        tableName = row.getCell(2).getStringCellValue();
                        tableDes = row.getCell(4).getStringCellValue();
                        break;
                    }
                }
                if (c == 2) {
                    cellValueMap.put("type", cellValue);
                }
                if (c == 3) {
                    cellValueMap.put("length", cellValue);
                }
                if (c == 4) {
                    cellValueMap.put("canBeNull", cellValue);
                }
                if (c == 5) {
                    cellValueMap.put("defValue", cellValue);
                }
                if (c == 6) {
                    cellValueMap.put("des", cellValue);
                    cellMap.put(mapKey, cellValueMap);
                }
            }
        }
        createGroovyString(cellMap, listMap, primaryKey, tableName, tableDes, email);
    }

    private static void index(String mapKey, Map<String, List<String>> listMap, Row row) {
        int cellCount = row.getPhysicalNumberOfCells();
        List<String> stringList = listMap.get(mapKey);
        if (stringList == null) {
            stringList = new ArrayList<>();
        }
        // 索引的起始位置是2
        int star = 2;
        StringBuilder key = null;
        for (int c = star; c < cellCount; c++) {
            String keyValue = row.getCell(c).getStringCellValue();
            if ("".equals(keyValue) || keyValue == null) {
                break;
            }
            if (key == null) {
                key = new StringBuilder(keyValue);
            } else {
                key.append(",").append(keyValue);
            }
        }
        stringList.add(Objects.requireNonNull(key).toString());
        listMap.put(mapKey, stringList);
    }

    private static void createGroovyString(Map<String, Map<String, String>> cellMap,
                                           Map<String, List<String>> listMap,
                                           String primaryKey,
                                           String tableName,
                                           String tableDes, String email) throws IOException {
        StringBuilder builder = new StringBuilder();
        tableHeader(tableName, tableDes, builder, email);
        if (!tableName.toLowerCase().contains("_tl")) {
            // 不是多语言表
            baseTable(builder);
        }
        // 字段
        column(primaryKey, cellMap, builder);
        // 索引
        createIndex(builder, listMap, tableName);
        changeLogEnd(builder);
        String path = "./groovy/" + tableName.toLowerCase() + ".groovy";
        File file = new File(path);
        file.createNewFile();
        OutputStreamWriter fw = new OutputStreamWriter(new FileOutputStream(file), "UTF-8");
        BufferedWriter bw = new BufferedWriter(fw);
        bw.write(builder.toString());
        bw.flush();
        bw.close();
        fw.close();
    }

    private static void tableHeader(String tableName, String description, StringBuilder builder, String email) {
        LocalDate localDate = LocalDate.now();
        builder.append("databaseChangeLog(logicalFilePath: 'script/db/").append(tableName.toLowerCase())
                .append(".groovy') {\n ")
                .append("   changeSet(author: \"").append(email).append("\", id: \"").append(localDate).append("_")
                .append(tableName.toLowerCase()).append("\") {\n")
                .append("def weight = 1\n").append(
                "        if (helper.isSqlServer()) {\n").append(
                "            weight = 2\n").append(
                "        } else if (helper.isOracle()) {\n").append(
                "            weight = 3\n").append(
                "        }\n").append(
                "        if (helper.dbType().isSupportSequence()) {\n").append(
                "            createSequence(sequenceName: '").append(tableName.toLowerCase())
                .append("', startValue: \"1\")\n").append(
                "        }\n").append(
                "        createTable(tableName: \"").append(tableName.toLowerCase())
                .append("\", remarks: \"")
                .append(description).append("\") {\n");
    }

    private static void baseTable(StringBuilder builder) {
        builder.append("            column(name: \"CREATED_BY\", type: \"bigint(20)\", defaultValue: \"1\", remarks: \"创建人\") {\n").append(
                "                constraints(nullable: \"false\")\n").append(
                "            }\n").append(
                "            column(name: \"CREATION_DATE\", type: \"datetime\", defaultValueComputed: \"CURRENT_TIMESTAMP\", remarks: \"创建时间\") {\n").append(
                "                constraints(nullable: \"false\")\n").append(
                "            }\n").append(
                "            column(name: \"LAST_UPDATED_BY\", type: \"bigint(20)\", defaultValue: \"1\", remarks: \"最后更新人\") {\n").append(
                "                constraints(nullable: \"false\")\n").append(
                "            }\n").append(
                "            column(name: \"LAST_UPDATE_DATE\", type: \"datetime\", defaultValueComputed: \"CURRENT_TIMESTAMP\", remarks: \"最后更新时间\") {\n").append(
                "                constraints(nullable: \"false\")\n").append(
                "            }\n").append(
                "            column(name: \"OBJECT_VERSION_NUMBER\", type: \"bigint(20)\", defaultValue: \"1\", remarks: \"版本号\") {\n").append(
                "                constraints(nullable: \"false\")\n").append(
                "            }\n");
    }

    private static void column(String primaryKey, Map<String, Map<String, String>> cellMap, StringBuilder builder) {
        cellMap.forEach((key, value) -> {
            // 判断是不是主键
            if (key.equals(primaryKey)) {
                builderString(value, builder, true, key);
            } else {
                builderString(value, builder, false, key);
            }
        });
        // 普通字段结束
        builder.append(
                "\n").append(
                "        }\n");
    }

    private static void builderString(Map<String, String> map, StringBuilder builder,
                                      boolean isPrimaryKey, String columnName) {

        String type = map.get("type");
        String length = map.get("length");
        String canBeNull = map.get("canBeNull");
        String defValue = map.get("defValue");
        String des = map.get("des");
        builder.append("            column(name: \"").append(columnName).append("\", type: \"");
        if ("BigInt".equals(type)) {
            builder.append("bigint(20)\"");
        } else if ("Varchar".equals(type)) {
            if (length != null && !"".equals(length)) {
                builder.append("varchar(\" + ").append(length).append(" * weight + \")\"");
            } else {
                builder.append("\"varchar\"");
            }
        } else {
            builder.append(type);
            if (length != null && !"".equals(length)) {
                builder.append("(").append(length).append(")\"");
            } else {
                builder.append("\"");
            }
        }
        if (defValue != null && !"".equals(defValue)) {
            builder.append(", defaultValue: \"").append(defValue).append("\"");
        }
        if (des != null && !"".equals(des)) {
            builder.append(", remarks: \"").append(des).append("\") {");
        }
        if (isPrimaryKey) {
            builder.append("\n").append(
                    "                constraints(primaryKey: \"true\", nullable: \"false\")\n").append(
                    "            }");
        } else {
            if ("是".equals(canBeNull)) {
                builder.append("\n").append(
                        "                constraints(nullable: \"true\")\n").append(
                        "            }");
            } else {
                builder.append("\n").append(
                        "                constraints(nullable: \"false\")\n").append(
                        "            }");
            }
        }
        builder.append("\n");
    }

    private static void createIndex(StringBuilder builder, Map<String, List<String>> mapMap, String tableName) {
        if (mapMap != null && mapMap.size() > 0) {
            // 获取唯一索引
            List<String> uniqueList = mapMap.get("唯一性索引");
            if (uniqueList != null && uniqueList.size() > 0) {
                uniqueList.forEach(a -> builder.append("        addUniqueConstraint(columnNames:").append("\"").append(a).append("\", tableName: \"")
                        .append(tableName).append("\", constraintName: \"请输入唯一索引名称\")\n"));
            }
            List<String> indexList = mapMap.get("普通索引");
            if (indexList != null && indexList.size() > 0) {
                indexList.forEach(a -> {
                    builder.append("        createIndex(tableName: \"").append(tableName)
                            .append("\", indexName: \"请输入普通索引名称\") {\n");
                    List<String> columns = Arrays.asList(a.split(","));
                    columns.forEach(s -> {
                        builder.append("            column(name: \"").append(s).append("\")\n");
                    });
                    builder.append("        }\n");
                });
            }
        }
    }


    private static void changeLogEnd(StringBuilder builder) {
        // 包含了changeSet
        builder.append("    }\n").append("}");
    }
}
