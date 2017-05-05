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

    private Validator<Sheet, Boolean> sheetBooleanValidator = (sheet -> sheet != null && sheet.getLastRowNum() >= 0);

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
    public static <T> OutputStream write(Class<T> clazz, List<T> rows, OutputStream os) throws IOException {
        return ExcelResolverFactory.singleSheetResolver.write(clazz, rows, os);
    }

    /**
     * 从输入流中读取excel数据
     *
     * @param is excel inputStream
     * @param clazz 行对象结构
     * @param <T> 行对象泛型
     * @return 数据对象集合
     * @throws InvalidFileException 文件校验异常
     */
    public <T> List<T> read(InputStream is, Class<T> clazz) throws InvalidFileException {
        return read(is, clazz, sheetBooleanValidator);
    }

    /**
     * 从输入流中读取指定名称的sheet内容到集合中
     * @param is 文件输入流
     * @param sheetName sheet名称
     * @param clazz 行对象结构
     * @param <T> 行对象泛型
     * @return 数据对象集合
     * @throws InvalidFileException 文件校验异常
     */
    public <T> List<T> read(InputStream is, String sheetName, Class<T> clazz) throws InvalidFileException {
        return read(is, sheetName, clazz, sheetBooleanValidator);
    }

    /**
     * 从输入流中读取指定序号的sheet内容到集合中
     * @param is 文件输入流
     * @param sheetIndex sheet序号,其中第一个sheet序号为0
     * @param clazz 行对象结构
     * @param <T> 行对象泛型
     * @return 数据对象集合
     * @throws InvalidFileException 文件校验异常
     */
    public <T> List<T> read(InputStream is, int sheetIndex, Class<T> clazz)
            throws InvalidFileException {
        return read(is, sheetIndex, clazz, sheetBooleanValidator);
    }

    /**
     * 从输入流中读取excel数据
     *
     * @param is excel inputStream
     * @param clazz 行对象结构
     * @param <T> 行对象泛型
     * @param sheetValidator sheet内容验证器
     * @return 数据对象集合
     * @throws InvalidFileException 文件校验异常
     */
    public <T> List<T> read(InputStream is, Class<T> clazz, Validator<Sheet, Boolean> sheetValidator)
            throws InvalidFileException {
        return ExcelResolverFactory.singleSheetResolver.read(is, clazz, sheetValidator);
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
        return ExcelResolverFactory.singleSheetResolver.read(is, sheetName, clazz, sheetValidator);
    }

    /**
     * 从输入流中读取指定序号的sheet内容到集合中
     * @param is 文件输入流
     * @param sheetIndex sheet序号,其中第一个sheet序号为0
     * @param clazz 行对象结构
     * @param sheetValidator sheet内容验证器
     * @param <T> 行对象泛型
     * @return 数据对象集合
     * @throws InvalidFileException 文件校验异常
     */
    public <T> List<T> read(InputStream is, int sheetIndex, Class<T> clazz, Validator<Sheet, Boolean> sheetValidator)
            throws InvalidFileException {
        return ExcelResolverFactory.singleSheetResolver.read(is, sheetIndex, clazz, sheetValidator);
    }

    private static class ExcelResolverFactory {
        private static SingleSheetExcelResolver singleSheetResolver = new SingleSheetExcelResolver();

    }
}
