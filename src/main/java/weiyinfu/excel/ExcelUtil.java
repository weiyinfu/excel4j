package weiyinfu.excel;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import weiyinfu.util.bean.BeanGs;
import weiyinfu.util.bean.GetterAndSetter;
import weiyinfu.util.bean.Gs;
import weiyinfu.util.bean.MapGs;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel工具类，用于Excel导入导出
 */
public class ExcelUtil {

//设置excel单元格的样式
private static HSSFCellStyle exportStyle(HSSFWorkbook workbook) {
    HSSFCellStyle cellStyle = workbook.createCellStyle();
    cellStyle.setAlignment(HorizontalAlignment.LEFT);
    cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    HSSFFont font = workbook.createFont();
    font.setFontName("微软雅黑");
    font.setFontHeightInPoints((short) 18);
    cellStyle.setFont(font);
    return cellStyle;
}

public static HSSFWorkbook exportBigExcel(String[] headers, String[] fields, ItemGetter getter) throws IOException {
    if (headers.length != fields.length) {
        throw new RuntimeException("headers的长度不等于fields的长度");
    }
    HSSFWorkbook workbook = new HSSFWorkbook();
    HSSFSheet sheet = workbook.createSheet("第一个sheet");
    int[] columnWidth = new int[headers.length];

    HSSFRow row = sheet.createRow(0);

    for (int i = 0; i < headers.length; i++) {
        HSSFCell cell = row.createCell(i);
        cell.setCellValue(headers[i]);
    }
    int i = 0;
    while (getter.next()) {
        i++;
        row = sheet.createRow(i);
        for (int j = 0; j < fields.length; j++) {
            HSSFCell cell = row.createCell(j);
            Object cellValue = getter.get(fields[j]);
            if (cellValue instanceof Integer || cellValue instanceof Double) {
                cell.setCellValue(Double.parseDouble(cellValue.toString()));
            } else {
                if (cellValue == null) cellValue = "";
                cell.setCellValue(cellValue.toString());
            }
            columnWidth[j] = Math.max(columnWidth[j], cellValue.toString().length());
        }
    }
    changeStyle(workbook, sheet, columnWidth);
    workbook.close();
    return workbook;
}

/**
 * @param headers excel标头名称
 * @param fields  成员变量名称
 * @param data    对象列表，对象可以是JavaBean也可以是Map
 */
public static HSSFWorkbook exportExcel(String[] headers, String[] fields, List<?> data) throws IOException {
    ListItemGetter getter = new ListItemGetter(data);
    return exportBigExcel(headers, fields, getter);
}

static void changeStyle(HSSFWorkbook workbook, HSSFSheet sheet, int columnWidth[]) {
    HSSFCellStyle style = exportStyle(workbook);
    for (int i = 0; i < columnWidth.length; i++) {
        sheet.setDefaultColumnStyle(i, style);
        int w = columnWidth[i];
        sheet.setColumnWidth(i, Math.min(40, w * 2 + 4) * 256);
    }
}

/**
 * 导入Excel
 *
 * @param cin     输入流
 * @param headers 需要的excel的表头列的名称
 * @param fields  对应的成员变量名
 * @param cls     返回列表中的元素类型
 */
public static <T> List<T> importExcel(InputStream cin, String[] headers, String[] fields, Class<T> cls) throws IOException, IllegalAccessException, InstantiationException, NoSuchMethodException, InvocationTargetException, NoSuchFieldException, InvalidFormatException {
    ListItemHandler<T> listHandler = new ListItemHandler<T>();
    importBigExcel(cin, headers, fields, cls, listHandler);
    return listHandler.a;
}

/**
 * 导入比较大的Excel文件，不需要全部读入
 */
public static <T> void importBigExcel(InputStream cin, String[] headers, String[] fields, Class<T> cls, ItemHandler<T> handler) throws IOException, InvalidFormatException, NoSuchFieldException, IllegalAccessException, InstantiationException, NoSuchMethodException, InvocationTargetException {
    if (headers.length != fields.length) {
        throw new RuntimeException("headers的长度不等于fields的长度");
    }


    Workbook workbook = WorkbookFactory.create(cin);
    Sheet sheet = workbook.getSheetAt(0);
    int col[] = new int[fields.length];
    Class<?>[] colTypes = new Class<?>[fields.length];
    Row row = sheet.getRow(0);
    HashMap<String, Integer> headerToCol = new HashMap<String, Integer>();
    for (int i = 0; i < headers.length; i++) {
        headerToCol.put(headers[i], i);
        colTypes[i] = cls.getDeclaredField(fields[i]).getType();
    }
    for (int i = 0; i < row.getLastCellNum(); i++) {
        Cell cell = row.getCell(i);
        String header = cell.getStringCellValue();
        if (headerToCol.containsKey(header)) {
            col[headerToCol.get(header)] = i;
        }
    }
    for (int i = 0; i < fields.length; i++) {
        if (col[i] == -1) {
            throw new RuntimeException("excel中缺少属性" + fields[i]);
        }
    }
    int rownum = 1;
    while (rownum <= sheet.getLastRowNum()) {
        row = sheet.getRow(rownum);
        T obj = cls.getDeclaredConstructor().newInstance();
        GetterAndSetter getterAndSetter = cls.isAssignableFrom(Map.class) ? new MapGs((Map<String, Object>) obj, true) : new BeanGs(obj, true);
        for (int i = 0; i < fields.length; i++) {
            try {
                Object val = null;
                Cell cell = row.getCell(col[i]);
                if (cell == null) continue;
                cell.setCellType(CellType.STRING);
                val = cell.getStringCellValue();
                val = Gs.seven(val, colTypes[i]);
                if (val == null) {
                    throw new RuntimeException("没法处理" + colTypes[i] + " 类型");
                }
                getterAndSetter.set(fields[i], val);
            } catch (NullPointerException e) {
                throw new RuntimeException(fields[i] + " not found", e);
            }
        }
        rownum++;
        handler.handle(obj);
    }
}


}