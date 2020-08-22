import com.alibaba.fastjson.JSON;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;
import weiyinfu.excel4j.ExcelUtil;
import cn.weiyinfu.gs.BeanGs;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

public class TestExcelUtil {

public static class User {
    String name;
    int age;

    public User() {
    }

    public User(String name, int age) {
        this.name = name;
        this.age = age;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    @Override
    public String toString() {
        return String.format("User(name=%s,age=%s)", name, age);
    }
}

//获取一个列表，用于调试
List<User> getList() {
    List<User> ans = new ArrayList<User>();
    for (int i = 0; i < 10; i++) {
        User haha = new User();
        haha.setName("name" + i);
        haha.setAge(i);
        ans.add(haha);
    }
    return ans;
}

@Test
public void testConstructor() throws NoSuchMethodException {
    Constructor<?>[] cons = User.class.getConstructors();
    Constructor<?>[] cons2 = User.class.getDeclaredConstructors();
    System.out.println(cons.length);
    System.out.println(cons2[0]);
    System.out.println(cons2[1]);
    System.out.println(cons2.length);
    Constructor<User> res = User.class.getDeclaredConstructor();
    System.out.println(res);
}

@Test
public void testBeanGs() {
    User u = new User("weiyinfu", 18);
    BeanGs gs = new BeanGs(u, true);
    System.out.println(gs.get("name"));
    System.out.println(gs.get("age"));
}

//测试java bean导出excel
@Test
public void testBean() throws IOException, NoSuchMethodException, IllegalAccessException, InvocationTargetException, InvalidFormatException, NoSuchFieldException, InstantiationException {
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

//测试map导出excel
@Test
public void testMap() throws NoSuchMethodException, IOException, IllegalAccessException, InvocationTargetException, InvalidFormatException, NoSuchFieldException, InstantiationException {
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
}
