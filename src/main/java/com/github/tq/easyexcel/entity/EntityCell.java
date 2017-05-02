package com.github.tq.easyexcel.entity;

import lombok.Data;
import lombok.experimental.Builder;
import org.apache.poi.ss.usermodel.CellType;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

/**
 * Created by tianque on 2017/4/18.
 */
@Data
@Builder
public class EntityCell {

    private int colNum;

    private CellType cellType;

    private String fieldName;

    private Field field;

    private String title;

}
