package weiyinfu.excel4j;

import cn.weiyinfu.gs.BeanGs;
import cn.weiyinfu.gs.GetterAndSetter;
import cn.weiyinfu.gs.MapGs;

import java.util.List;
import java.util.Map;

public class ListItemGetter implements ItemGetter {
List<?> data;
int index;
GetterAndSetter gs;

public ListItemGetter(List<?> data) {
    this.data = data;
    index = -1;
}

@Override
public boolean next() {
    index++;
    if (index >= data.size()) {
        return false;
    } else {
        Object obj = data.get(index);
        if (obj.getClass() == Map.class) {
            Map<String, Object> ma = (Map<String, Object>) obj;
            this.gs = new MapGs(ma, true);
        } else {
            this.gs = new BeanGs(obj, true);
        }
        return true;
    }
}


@Override
public Object get(String attr) {
    return gs.get(attr);
}

@Override
public void set(String attr, Object valueObj) {
    throw new RuntimeException("set函数无法调用");
}
}
