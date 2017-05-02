package com.github.tq.easyexcel.helper;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import com.github.tq.easyexcel.entity.EntityCell;

/**
 * Created by tianque on 2017/5/2.
 */
public abstract class ExcelResolver {

    Object getCellValue(Cell col, EntityCell cell) {
        String strCellValue;
        switch (col.getCellTypeEnum()) {
            case STRING:
                strCellValue = col.getStringCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(col)) {
                    // 如果是date类型则 ，获取该cell的date值
                    strCellValue = DateFormatUtils.ISO_DATETIME_FORMAT
                            .format(DateUtil.getJavaDate(col.getNumericCellValue()));
                } else { // 纯数字
                    col.setCellType(CellType.STRING);
                    strCellValue = col.getStringCellValue();
                }
                break;
            case BOOLEAN:
                strCellValue = String.valueOf(col.getBooleanCellValue());
                break;
            case BLANK:
                strCellValue = "";
                break;
            default:
                strCellValue = "";
                break;
        }

        return changeCellValueType(cell, strCellValue);
    }

    private Object changeCellValueType(EntityCell cell, String value) {
        Class<?> type = cell.getField().getType();
        if (type == String.class) {
            return value;
        }
        if (type.isPrimitive()) {
            if (type == short.class) {
                return Short.valueOf(value).shortValue();
            } else if (type == int.class) {
                return Integer.valueOf(StringUtils.substringBefore(value, ".")).intValue();
            } else if (type == long.class) {
                return Long.parseLong(StringUtils.substringBefore(value, "."));
            } else if (type == float.class) {
                return Float.valueOf(value).floatValue();
            } else if (type == double.class) {
                return Double.valueOf(value).doubleValue();
            } else if (type == byte.class) {
                return Byte.valueOf(value);
            } else if (type == boolean.class) {
                return Boolean.valueOf(value).booleanValue();
            } else if (type == char.class) {
                return value.charAt(0);
            }
        }

        if (type == Short.class) {
            return Short.valueOf(value);
        } else if (type == Integer.class) {
            return Integer.valueOf(value);
        } else if (type == Long.class) {
            return Long.valueOf(value);
        } else if (type == Float.class) {
            return Float.valueOf(value);
        } else if (type == Double.class) {
            return Double.parseDouble(value);
        } else if (type == Date.class) {
            try {
                return DateFormatUtils.ISO_DATETIME_FORMAT.parse(value);
            } catch (ParseException e) {
                String pattern = cell.getField().getAnnotation(com.github.tq.easyexcel.annotation.Cell.class)
                        .dateFormatPattern();
                try {
                    return new SimpleDateFormat(pattern).parse(value);
                } catch (ParseException e1) {
                    throw new IllegalArgumentException("can not parse date from " + value);
                }
            }
        } else if (type == Boolean.class) {
            return Boolean.valueOf(value);
        }
        return value;
    }

    void setCellValue(Cell cell, EntityCell entityCell, Object cellVal) {
        switch (entityCell.getCellType()) {
            case STRING:
                cell.setCellValue(String.valueOf(cellVal));
                break;
            case BOOLEAN:
                cell.setCellValue(Boolean.valueOf(String.valueOf(cellVal)));
                break;
            case NUMERIC:
                cell.setCellValue(String.valueOf(cellVal));
                break;
            case FORMULA:
                break;
            case _NONE:
                break;
            case BLANK:
                Class<?> fieldType = entityCell.getField().getType();
                if (fieldType == Date.class) {
                    cell.setCellType(CellType.STRING);
                    String pattern = entityCell.getField().getAnnotation(com.github.tq.easyexcel.annotation.Cell.class)
                            .dateFormatPattern();
                    String strVal = DateFormatUtils.format((Date) cellVal, pattern);
                    cell.setCellValue(strVal);
                } else if (fieldType == Calendar.class) {
                    cell.setCellValue((Calendar) cellVal);
                } else if (isNumeric(fieldType)) {
                    cell.setCellValue(Double.valueOf(String.valueOf(cellVal)));
                } else if (fieldType == String.class) {
                    cell.setCellValue(String.valueOf(cellVal));
                } else if (fieldType == boolean.class || fieldType == Boolean.class) {
                    cell.setCellValue(Boolean.valueOf(String.valueOf(cellVal)));
                } else {
                    cell.setCellValue(String.valueOf(cellVal));
                }
                break;
            default:
                break;
        }
    }

    private boolean isNumeric(Class<?> type) {
        if (type.isPrimitive() && (type == short.class || type == int.class || type == long.class || type == float.class
                || type == double.class)) {
            return true;
        } else if (type == Short.class || type == Integer.class || type == Long.class || type == Float.class
                || type == Double.class) {
            return true;
        }

        return false;
    }
}
