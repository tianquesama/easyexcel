package com.github.tq.easyexcel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.Date;

import com.github.tq.easyexcel.helper.SingleSheetExcelResolver;

/**
 * Created by nijun on 2017/4/25.
 */
public class TestWrite {

    public static void main(String[] args) throws IOException {

        SingleSheetExcelResolver xlsHelper = new SingleSheetExcelResolver();

        XlsBean bean = new XlsBean();
        bean.setSs("aaaa");
        bean.setB(false);
        bean.setDate(new Date());
        bean.setI(1);
        bean.setL(222L);

        FileOutputStream fos = new FileOutputStream("./target/a.xlsx");
        xlsHelper.write(XlsBean.class, Collections.singletonList(bean), fos);
    }

}
