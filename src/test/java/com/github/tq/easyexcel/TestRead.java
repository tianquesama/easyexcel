package com.github.tq.easyexcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import com.github.tq.easyexcel.exception.InvalidFileException;
import com.github.tq.easyexcel.helper.SingleSheetExcelResolver;

/**
 * Created by nijun on 2017/4/25.
 */
public class TestRead {

    public static void main(String[] args) throws IOException {

        SingleSheetExcelResolver xlsHelper = new SingleSheetExcelResolver();

        FileInputStream fis = new FileInputStream("./target/a.xlsx");
        try {
            List<XlsBean> list = xlsHelper.read(fis, XlsBean.class, (rows -> rows.getLastRowNum() >= 1));
            System.out.println(list.size());
        } catch (InvalidFileException e) {
            e.printStackTrace();
        }
    }

}
