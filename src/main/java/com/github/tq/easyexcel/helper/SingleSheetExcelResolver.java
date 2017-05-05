package com.github.tq.easyexcel.helper;

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.commons.lang3.time.FastDateFormat;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.github.tq.easyexcel.entity.EntityCell;
import com.github.tq.easyexcel.entity.EntitySheet;
import com.github.tq.easyexcel.exception.InvalidBeanException;
import com.github.tq.easyexcel.exception.InvalidFileException;
import com.github.tq.easyexcel.validator.Validator;

/**
 * Created by tianque on 2017/4/18.
 */
public class SingleSheetExcelResolver {

    private SheetResolver resolver = new SheetResolver();

    private FastDateFormat dateFormat = DateFormatUtils.ISO_DATETIME_FORMAT;

    /**
     * 将数据写到输出流中
     *
     * @param clazz 行数据结构
     * @param rows excel中的行数据
     * @param os 输出流
     * @param <T> 行对象泛型
     * @return 返回输出流
     * @throws IOException IO异常
     */
    public <T> OutputStream write(Class<T> clazz, List<T> rows, OutputStream os) throws IOException {
        EntitySheet entitySheet = resolver.resolve(clazz);

        SXSSFWorkbook wb = new SXSSFWorkbook();
        SXSSFSheet sheet = wb.createSheet("sheet1");
        SXSSFRow titleRow = sheet.createRow(0);

        entitySheet.getCells().forEach(cell -> titleRow.createCell(cell.getColNum(), CellType.STRING).setCellValue(cell.getTitle()));

        int rownum = 1;
        for (T rowBean : rows) {
            SXSSFRow row = sheet.createRow(rownum);
            for (EntityCell entityCell : entitySheet.getCells()) {
                Object cellVal;
                try {
                    cellVal = PropertyUtils.getProperty(rowBean, entityCell.getFieldName());
                } catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
                    throw new RuntimeException("can not access value of field ".concat(entityCell.getFieldName()));
                }

                SXSSFCell rowCell = row.createCell(entityCell.getColNum(), entityCell.getCellType());
                if (cellVal != null) {
                    setCellValue(rowCell, entityCell, cellVal);
                }
            }

            rownum++;
        }

        if (os == null) {
            os = new ByteArrayOutputStream();
        }
        wb.write(os);
        return os;
    }

    /**
     * 从输入流中读取指定名称的sheet内容到集合中
     * @param is 文件输入流
     * @param sheetName sheet名称
     * @param clazz 行对象结构
     * @param sheetValidator sheet内容验证器
     * @param <T> 行对象泛型
     * @return 数据对象集合
     * @throws InvalidFileException 文件校验异常
     */
    public <T> List<T> read(InputStream is, String sheetName, Class<T> clazz, Validator<Sheet, Boolean> sheetValidator)
            throws InvalidFileException {
        Workbook wb = null;
        try {
            wb = WorkbookFactory.create(is);
        } catch (IOException | InvalidFormatException e) {
            throw new InvalidFileException("inputStream is not a excel file", e);
        }

        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) {
            throw new InvalidFileException("can not found the sheet named ".concat(sheetName));
        }
        return read(sheet, clazz, sheetValidator);
    }

    public <T> List<T> read(InputStream is, Class<T> clazz, Validator<Sheet, Boolean> sheetValidator)
            throws InvalidFileException {
        return read(is, 0, clazz, sheetValidator);
    }

    public <T> List<T> read(InputStream is, int sheetIndex, Class<T> clazz, Validator<Sheet, Boolean> sheetValidator)
            throws InvalidFileException {

        if (is == null || clazz == null) {
            throw new NullPointerException("inputStream or clazz is null");
        }

        Workbook wb = null;
        try {
            wb = WorkbookFactory.create(is);
        } catch (IOException | InvalidFormatException e) {
            throw new InvalidFileException("inputStream is not a excel file", e);
        }

        Sheet sheet = wb.getSheetAt(sheetIndex);

        return read(sheet, clazz, sheetValidator);
    }

    private <T> List<T> read(Sheet sheet, Class<T> clazz, Validator<Sheet, Boolean> sheetValidator) throws InvalidFileException {
        if (sheet == null) {
            throw new InvalidFileException("can not found any sheet.");
        }

        if (sheetValidator != null && !sheetValidator.apply(sheet)) {
            throw new InvalidFileException("failed to access the file validation");
        }

        EntitySheet entitySheet = resolver.resolve(clazz);
        List<EntityCell> cells = entitySheet.getCells();

        int rowSize = sheet.getLastRowNum();

        List<T> beanList = new ArrayList<>();
        for (int i = sheet.getFirstRowNum() + 1; i <= rowSize; i++) {
            Row row = sheet.getRow(i);
            try {
                T clazzBean = clazz.newInstance();
                boolean emptyRow = true;
                for (EntityCell cell : cells) {
                    String fieldName = cell.getFieldName();
                    int colNum = cell.getColNum();
                    Cell col = row.getCell(colNum, RETURN_BLANK_AS_NULL);
                    if (col == null) {
                        continue;
                    }
                    Object val = getCellValue(col, cell);
                    if (val != null && !"".equals(val)) {
                        try {
                            PropertyUtils.setProperty(clazzBean, fieldName, val);
                            emptyRow = false;
                        } catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
                            throw new InvalidBeanException("can not found a setter for field :".concat(fieldName), e);
                        }
                    }
                }
                if (emptyRow) {
                    continue;
                }
                beanList.add(clazzBean);
            } catch (InstantiationException | IllegalAccessException e) {
                throw new InvalidBeanException("can not access or found a no arg constructor of ".concat(clazz.getSimpleName()), e);
            }
        }

        return beanList;
    }


    private Object getCellValue(Cell col, EntityCell cell) {
        String strCellValue;
        switch (col.getCellTypeEnum()) {
            case STRING:
                strCellValue = col.getStringCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(col)) {
                    // 如果是date类型则 ，获取该cell的date值
                    strCellValue = dateFormat.format(DateUtil.getJavaDate(col.getNumericCellValue()));
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

        if (StringUtils.isEmpty(strCellValue)) {
            return strCellValue;
        }

        try {
            return changeCellValueType(cell, strCellValue);
        } catch (NumberFormatException e) {
            String strErrorVal = cell.getField().getAnnotation(com.github.tq.easyexcel.annotation.Cell.class).errorValue();
            return changeCellValueType(cell, strErrorVal);
        }

    }

    private Object changeCellValueType(EntityCell cell, String value) {
        Class<?> type = cell.getField().getType();
        if (type == String.class) {
            return value;
        }
        if (type.isPrimitive()) {
            if (type == short.class) {
                return Short.valueOf(StringUtils.substringBefore(value, ".")).shortValue();
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
                char[] charArray = value.toCharArray();
                if (charArray.length <= 0) {
                    return "";
                }
                return charArray[0];
            }
        }

        if (type == Short.class) {
            return Short.valueOf(StringUtils.substringBefore(value, "."));
        } else if (type == Integer.class) {
            return Integer.valueOf(StringUtils.substringBefore(value, "."));
        } else if (type == Long.class) {
            return Long.valueOf(StringUtils.substringBefore(value, "."));
        } else if (type == Float.class) {
            return Float.valueOf(value);
        } else if (type == Double.class) {
            return Double.parseDouble(value);
        } else if (type == Date.class) {
            try {
                return dateFormat.parse(value);
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
        } else if (type == BigDecimal.class) {
            return BigDecimal.valueOf(Double.valueOf(value));
        } else if (type == BigInteger.class) {
            return BigInteger.valueOf(Long.valueOf(value));
        }
        return value;
    }

    private void setCellValue(Cell cell, EntityCell entityCell, Object cellVal) {
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
