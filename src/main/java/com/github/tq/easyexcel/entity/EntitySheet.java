package com.github.tq.easyexcel.entity;

import lombok.Data;
import lombok.experimental.Builder;

import java.util.List;

/**
 * Created by nijun on 2017/4/18.
 */
@Data
@Builder
public class EntitySheet {

    private String sheetName = "sheet";

    private List<EntityCell> cells;

    //TODO 还应包含列格式等信息
}
