package com.demo.excle;

import lombok.Data;
import lombok.ToString;
import lombok.experimental.Accessors;

import java.util.Date;
import java.util.Map;


@Data
@Accessors(chain = true)
@ToString
public class AbstractModel {
    // 主题名称
    private String subject;

    //身份表示 1 园区 2 非园区 3 园区员工 4 非园区员工
    private Integer flag;

    //品种结果集映射
    private Map<String, Object> resultMap;

    //当前时间
    private Date currentDate;

    //当前表索引 对应  周几
    private Integer week;
}
