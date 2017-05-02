package com.github.tq.easyexcel.helper;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;

import com.github.tq.easyexcel.exception.InvalidFileException;
import com.github.tq.easyexcel.validator.Validator;

/**
 * Created by tianque on 2017/5/2.
 */
public class ExcelHelper {

    private Validator<Sheet, Boolean> defaultValidator = (sheet -> sheet.getLastRowNum() >= 0);

    public static <T> OutputStream write(Class<T> clazz, List<T> rows, OutputStream os) throws IOException {
        return ExcelResolverFactory.singleSheetResolver.write(clazz, rows, os);
    }

    public <T> List<T> read(InputStream is, Class<T> clazz) throws InvalidFileException {
        return read(is, clazz, defaultValidator);
    }

    public <T> List<T> read(InputStream is, Class<T> clazz, Validator<Sheet, Boolean> sheetValidator)
            throws InvalidFileException {
        return ExcelResolverFactory.singleSheetResolver.read(is, clazz, sheetValidator);
    }

    private static class ExcelResolverFactory {
        private static SingleSheetExcelResolver singleSheetResolver = new SingleSheetExcelResolver();

    }
}
