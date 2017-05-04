package com.github.tq.easyexcel.helper;

import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellType;

import com.github.tq.easyexcel.annotation.Cell;
import com.github.tq.easyexcel.entity.EntityCell;

/**
 * Created by nijun on 2017/4/20.
 */
public class CellHelper {

    private Map<Class<?>, List<EntityCell>> cellMap = new ConcurrentHashMap<>();

    protected List<EntityCell> getCells(Class<?> clazz) {

        if (cellMap.containsKey(clazz)) {
            return cellMap.get(clazz);
        }

        Field[] fieldArray = clazz.getDeclaredFields();
        if (fieldArray == null || fieldArray.length <= 0) {
            throw new IllegalArgumentException("class ".concat(clazz.getName()).concat(" has no fields"));
        }

        List<Field> fields = Arrays.asList(fieldArray);
        fields.removeIf(field -> !field.isAnnotationPresent(Cell.class));

        if (fields.isEmpty()) {
            throw new IllegalArgumentException("class ".concat(clazz.getName()).concat(" has no cell fields"));
        }

        List<EntityCell> cells = new ArrayList<>();
        int colSize = fields.size();
        for (Field field : fields) {
            checkFieldType(field.getType());
            Cell anno = field.getAnnotation(Cell.class);
            int order = anno.order();
            if (order > colSize) {
                colSize = order;
            }
            String title = anno.title();
            if (StringUtils.isEmpty(title)) {
                title = field.getName();
            }
            CellType cellType = anno.type();
            if (cellType == CellType._NONE) {
                cellType = CellType.BLANK;
            }
            EntityCell cell = EntityCell.builder().cellType(cellType).title(title).colNum(order - 1).field(field)
                    .fieldName(field.getName()).build();
            cells.add(cell);
        }

        EntityCell[] cellArray = new EntityCell[colSize];
        for (EntityCell cell : cells) {
            int col = cell.getColNum();
            if (col < 0) {
                col = findFirstNullIndex((Object[]) cellArray);
                cell.setColNum(col);
            }

            if (cellArray[col] != null) {
                throw new IllegalArgumentException();
            }
            cellArray[col] = cell;
        }

        cells.sort(Comparator.comparingInt(EntityCell::getColNum));
        cellMap.put(clazz, cells);
        return cells;
    }

    private int findFirstNullIndex(Object... array) {
        if (array == null) {
            return -1;
        }
        for (int i = 0; i < array.length; i++) {
            if (array[i] == null) {
                return i;
            }
        }
        return array.length - 1;
    }

    private void checkFieldType(Class<?> type) {
        if (!type.isPrimitive() && type != Date.class && type != Calendar.class && type != Boolean.class
                && type != String.class && type != Integer.class && type != Short.class && type != Long.class
                && type != Double.class) {
            throw new IllegalArgumentException("unSupport field type " + type.getName());
        }
    }

}
