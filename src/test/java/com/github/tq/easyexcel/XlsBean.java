package com.github.tq.easyexcel;

import java.util.Date;

import com.github.tq.easyexcel.annotation.Cell;

/**
 * Created by nijun on 2017/4/26.
 */
public class XlsBean {
    @Cell(order = 8)
    private String ss;

    @Cell(dateFormatPattern = "yyyy-MM-dd")
    private Date date;

    @Cell(title = "啦啦啦啦")
    private Integer i;

    @Cell(errorValue = "88")
    private long l;

    @Cell
    private Boolean b;













    public String getSs() {
        return ss;
    }

    public void setSs(String ss) {
        this.ss = ss;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public int getI() {
        return i;
    }

    public void setI(int i) {
        this.i = i;
    }

    public long getL() {
        return l;
    }

    public void setL(long l) {
        this.l = l;
    }

    public boolean isB() {
        return b;
    }

    public void setB(boolean b) {
        this.b = b;
    }
}
