package com.github.tq.easyexcel.validator;

import com.github.tq.easyexcel.exception.InvalidFileException;

/**
 * Created by tianque on 2017/4/24.
 */
@FunctionalInterface
public interface Validator<T, R> {

    R apply(T t) throws InvalidFileException;
}
