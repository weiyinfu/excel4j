package excel;

import utils.bean.BeanGetterAndSetter;
import utils.bean.GetterAndSetter;
import utils.bean.MapGetterAndSetter;

import java.util.List;
import java.util.Map;

public class ListItemGetter implements ItemGetter {
List<?> data;
int index;
GetterAndSetter gs;

public ListItemGetter() {

}

public ListItemGetter(List<?> data) {
    init(data);
}

@Override
public boolean next() {
    index++;
    if (index >= data.size()) {
        return false;
    } else {
        gs.init(data.get(index));
        return true;
    }
}


@Override
public void init(Object obj) {
    this.data = (List<?>) obj;
    this.gs = data.get(0).getClass() == Map.class ? new MapGetterAndSetter() : new BeanGetterAndSetter();
    index = -1;
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
