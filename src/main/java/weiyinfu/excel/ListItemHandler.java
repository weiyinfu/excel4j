package weiyinfu.excel;

import java.util.ArrayList;
import java.util.List;

class ListItemHandler<T> implements ItemHandler<T> {
public List<T> a = new ArrayList<T>();

@Override
public void handle(T it) {
    a.add(it);
}
}