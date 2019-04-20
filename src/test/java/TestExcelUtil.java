import com.alibaba.fastjson.JSON;
import excel.ExcelUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import utils.User;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

public class TestExcelUtil {

//获取一个列表，用于调试
static List<User> getList() {
    List<User> ans = new ArrayList<User>();
    for (int i = 0; i < 10; i++) {
        User haha = new User();
        haha.setName("name" + i);
        haha.setAge(i);
        ans.add(haha);
    }
    return ans;
}

//测试java bean导出excel
static void testBean() throws IOException, NoSuchMethodException, IllegalAccessException, InvocationTargetException, InvalidFormatException, NoSuchFieldException, InstantiationException {
    ExcelUtil.exportExcel(new String[]{"姓名", "年龄"},
            new String[]{"name", "age"},
            getList()).write(new File("haha.xls"));
    List<User> a = ExcelUtil.importExcel(
            new FileInputStream("haha.xls"),
            new String[]{"姓名", "年龄"},
            new String[]{"name", "age"},
            User.class);
    //.out.println(JSON.toJSONString(a, true));
}

//测试map导出excel
static void testMap() throws NoSuchMethodException, IOException, IllegalAccessException, InvocationTargetException, InvalidFormatException, NoSuchFieldException, InstantiationException {
    ExcelUtil.exportExcel(new String[]{"姓名", "年龄"},
            new String[]{"name", "age"},
            getList()).write(new File("haha.xls"));
    List<User> a = ExcelUtil.importExcel(
            new FileInputStream("haha.xls"),
            new String[]{"姓名", "年龄"},
            new String[]{"name", "age"},
            User.class);
    System.out.println(JSON.toJSONString(a, true));
}

public static void main(String[] args) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException, IOException, NoSuchFieldException, InstantiationException, InvalidFormatException {
    testMap();
}
}
