package com.github.tq.easyexcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.tq.easyexcel.exception.InvalidFileException;
import com.github.tq.easyexcel.helper.SingleSheetExcelResolver;

/**
 * Created by nijun on 2017/4/25.
 */
public class TestHelper {
    private final Logger logger = LoggerFactory.getLogger(this.getClass());

    private SingleSheetExcelResolver xlsHelper;

    @Before
    public void setUp() throws Exception {
        xlsHelper = new SingleSheetExcelResolver();
    }

    @Test
    public void testWrite() throws Exception {

        XlsBean bean = new XlsBean();
        bean.setSs("aaaa");
        bean.setB(false);
        bean.setDate(new Date());
        bean.setI(1);
        bean.setL(222L);

        long start = System.currentTimeMillis();
        logger.debug("start time : {}", start);
        FileOutputStream fos = new FileOutputStream("./target/a.xlsx");
        xlsHelper.write(XlsBean.class, Collections.singletonList(bean), fos);
        logger.debug("============>>>>>>>>>>>>>>>> write cost time:{}", System.currentTimeMillis() - start);
        fos.close();
    }


    @Test
    public void testRead() throws Exception {

        FileInputStream fis = new FileInputStream("./target/a.xlsx");
        try {
            long start = System.currentTimeMillis();
            logger.debug("start time : {}", start);
            List<XlsBean> list = xlsHelper.read(fis, XlsBean.class, (rows -> rows.getLastRowNum() >= 1));
            logger.debug("============>>>>>>>>>>>>>>>> read cost time:{}", System.currentTimeMillis() - start);
            Assert.assertNotNull(list);
        } catch (InvalidFileException e) {
            Assert.assertTrue(true);
        }
    }

    @Test
    public void test3000Write() throws Exception {
        List<XlsBean> beanList = new ArrayList<>();
        for (int i = 0; i < 3000; i++) {
            XlsBean bean = new XlsBean();
            bean.setSs("aaaa");
            bean.setB(false);
            bean.setDate(new Date());
            bean.setI(1);
            bean.setL(222L);
            beanList.add(bean);
        }
        long start = System.currentTimeMillis();
        logger.debug("start time : {}", start);
        FileOutputStream fos = new FileOutputStream("./target/list3k.xlsx");
        xlsHelper.write(XlsBean.class, beanList, fos);
        logger.debug("============>>>>>>>>>>>>>>>> 3k write cost time:{}", System.currentTimeMillis() - start);
        fos.close();
    }

    @Test
    public void test30000Write() throws Exception {
        List<XlsBean> beanList = new ArrayList<>();
        for (int i = 0; i < 30000; i++) {
            XlsBean bean = new XlsBean();
            bean.setSs("aaaa");
            bean.setB(false);
            bean.setDate(new Date());
            bean.setI(1);
            bean.setL(222L);
            beanList.add(bean);
        }
        long start = System.currentTimeMillis();
        logger.debug("start time : {}", start);
        FileOutputStream fos = new FileOutputStream("./target/list3w.xlsx");
        xlsHelper.write(XlsBean.class, beanList, fos);
        logger.debug("============>>>>>>>>>>>>>>>> 3w write cost time:{}", System.currentTimeMillis() - start);
        fos.close();
    }

    @Test
    public void test3kRead() throws Exception {
        FileInputStream fis = new FileInputStream("./target/list3k.xlsx");
        try {
            long start = System.currentTimeMillis();
            logger.debug("start time : {}", start);
            List<XlsBean> list = xlsHelper.read(fis, XlsBean.class, (rows -> rows.getLastRowNum() >= 1));
            logger.debug("============>>>>>>>>>>>>>>>> 3k read cost time:{}", System.currentTimeMillis() - start);
            Assert.assertNotNull(list);
        } catch (InvalidFileException e) {
            Assert.assertTrue(true);
        }
    }

    @Test
    public void test3wRead() throws Exception {
        FileInputStream fis = new FileInputStream("./target/list3w.xlsx");
        try {
            long start = System.currentTimeMillis();
            logger.debug("start time : {}", start);
            List<XlsBean> list = xlsHelper.read(fis, XlsBean.class, (rows -> rows.getLastRowNum() >= 1));
            logger.debug("============>>>>>>>>>>>>>>>> 3w read cost time:{}", System.currentTimeMillis() - start);
            Assert.assertNotNull(list);
        } catch (InvalidFileException e) {
            Assert.assertTrue(true);
        }
    }
}
