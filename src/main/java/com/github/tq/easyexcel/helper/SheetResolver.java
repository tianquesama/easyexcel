package com.github.tq.easyexcel.helper;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;

import com.github.tq.easyexcel.annotation.Cell;
import com.github.tq.easyexcel.entity.EntityCell;
import com.github.tq.easyexcel.entity.EntitySheet;

/**
 * Created by tianque on 2017/4/18.
 */
public class SheetResolver {

    private EntitySheetHolder sheetHolder = new EntitySheetHolder();

    private CellHelper cellHelper = new CellHelper();

    protected EntitySheet resolve(Class<?> clazz) {
        if (sheetHolder.containsKey(clazz)) {
            return sheetHolder.get(clazz);
        }

        EntitySheet.EntitySheetBuilder sheetBuilder = EntitySheet.builder();

        Field[] fieldArray = clazz.getDeclaredFields();

        List<Field> fields = Arrays.asList(fieldArray);
        fields.removeIf(field -> !field.isAnnotationPresent(Cell.class));

        List<EntityCell> cells = cellHelper.getCells(clazz);

        sheetBuilder.cells(cells);

        EntitySheet sheet = sheetBuilder.build();
        sheetHolder.put(clazz, sheet);
        return sheet;
    }

}
