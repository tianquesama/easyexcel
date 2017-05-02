package com.github.tq.easyexcel.helper;

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.github.tq.easyexcel.entity.EntityCell;
import com.github.tq.easyexcel.entity.EntitySheet;
import com.github.tq.easyexcel.exception.InvalidBeanException;
import com.github.tq.easyexcel.exception.InvalidFileException;
import com.github.tq.easyexcel.validator.Validator;

/**
 * Created by tianque on 2017/4/18.
 */
public class SingleSheetExcelResolver extends ExcelResolver{

    private SheetResolver resolver = new SheetResolver();

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

    public <T> List<T> read(InputStream is, Class<T> clazz, Validator<Sheet, Boolean> sheetValidator)
            throws InvalidFileException {

        if (is == null || clazz == null) {
            throw new NullPointerException("inputStream or clazz is null");
        }

        Workbook wb;

        try {
            wb = new XSSFWorkbook(is);
        } catch (IOException e) {
            try {
                wb = new HSSFWorkbook(is);
            } catch (IOException e1) {
                throw new InvalidFileException("inputStream is not a excel file");
            }
        }

        Sheet sheet = wb.getSheetAt(0);

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
                cells.forEach(cell -> {
                    String fieldName = cell.getFieldName();
                    int colNum = cell.getColNum();
                    Cell col = row.getCell(colNum, RETURN_BLANK_AS_NULL);
                    Object val = getCellValue(col, cell);
                    try {
                        PropertyUtils.setProperty(clazzBean, fieldName, val);
                    } catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
                        throw new InvalidBeanException("can not found a setter for field :".concat(fieldName));
                    }
                });
                beanList.add(clazzBean);
            } catch (InstantiationException | IllegalAccessException e) {
                throw new InvalidBeanException("can not access or found a no arg constructor of ".concat(clazz.getSimpleName()));
            }
        }

        return beanList;
    }

}
