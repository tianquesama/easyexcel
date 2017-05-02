package com.github.tq.easyexcel.exception;

/**
 * Created by tianque on 2017/4/24.
 */
public class InvalidBeanException extends RuntimeException {

    public InvalidBeanException(String message) {
        super(message);
    }

    public InvalidBeanException(String message, Throwable cause) {
        super(message, cause);
    }

    public InvalidBeanException(Throwable cause) {
        super(cause);
    }

    public InvalidBeanException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }

}
