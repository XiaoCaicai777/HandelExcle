package com.demo.excle;


import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.io.IoUtil;
import cn.hutool.core.util.StrUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

/**
 *
 */
public class MainExcle {

    /**
     * 解析数据源
     * @return
     */
    public static Map<Integer, List<AbstractModel>> getSource() {
        String url = System.getProperty("user.dir") + "\\"  + "园所蔬菜请购计划样表.xls";
        // 园区幼儿
        List<Integer> children = Arrays.asList(4,5,6,8,9,10,11);
        // 园区员工
        List<Integer> employee = Arrays.asList(27,28,29,31,32,33,34);
        // 非园区幼儿
        List<Integer> noChildren = Arrays.asList(12,13,14,15,16,17,18,19,20,21,22,23,24);
        // 非园区幼儿 排除海伦幼儿
//        List<Integer> noChildren = Arrays.asList(12,13,14,15,16,17,19,20,21,22,23,24);
        // 非园区员工
        List<Integer> noEmployee = Arrays.asList(35,36,37,38,39,40,41,42,43,44,45,46,47,48);
        // 非园区员工  排除总部员工和海伦员工
//        List<Integer> noEmployee = Arrays.asList(35,36,37,38,39,40,42,43,45,46,47,48);
        List<AbstractModel> resultList = new ArrayList<>();
        try (InputStream is= getFileInputStream(url)) {
            try (BufferedInputStream bis = new BufferedInputStream(is)) {
                try (Workbook workbook = XSSFWorkbookFactory.create(bis)) {
                    // 循环表
                    for (int i = 1; i < workbook.getNumberOfSheets(); i++) {
                        Map<Integer, Object> headMap = new HashMap<>();;
                        Sheet sheetAt = workbook.getSheetAt(i);
                        Date currentDate = sheetAt.getRow(0).getCell(4).getDateCellValue();
                        // 循环表中的行
                        for (int j = 2; j < 49; j++) {
                            // 获取表头索引映射
                            int lastCellNum = (int) sheetAt.getRow(j).getLastCellNum();
                            if (j == 2) {
                                for (int v = 3; v < lastCellNum; v++) {
                                    String stringCellValue = sheetAt.getRow(j).getCell(v).getStringCellValue();
                                    if (StrUtil.isNotEmpty(stringCellValue)) {
                                        headMap.put(v, stringCellValue);
                                    }
                                }
                                continue;
                            }
                            //跳过第4， 24， 25行 2022-09-16 排除海论幼儿园和总部员工
                            if (j == 3 || j == 25 || j == 26 || j == 7 || j == 30 || j == 41 || j == 44 || j == 18) {
                                continue;
                            }
                            AbstractModel abstractModel = new AbstractModel();
                            // 保存结果集
                            Map<String, Object> map = new LinkedHashMap<>();
                            //当前行
                            Row row = sheetAt.getRow(j);
                            for (int k = 0; k < row.getLastCellNum(); k++) {
                                // 获取主题名
                                if (k == 0) {
                                    String stringCellValue = null;
                                    if (row.getCell(k) != null) {
                                        stringCellValue = row.getCell(k).getStringCellValue();
                                    }
                                    // 如果获取到的主题为空 跳出循环
                                    if (StrUtil.isEmpty(stringCellValue)) {
                                        break;
                                    }
                                    abstractModel.setSubject(stringCellValue);
                                    abstractModel.setWeek(i - 1);
                                    // 只保存第一个表中的日期
                                    if (i == 1) {
                                        abstractModel.setCurrentDate(currentDate);
                                    }
                                    // 设置身份标识
                                    if (children.contains(j)) {
                                        abstractModel.setFlag(1);
                                    } else if (noChildren.contains(j)) {
                                        abstractModel.setFlag(2);
                                    } else if (employee.contains(j)) {
                                        abstractModel.setFlag(3);
                                    } else if (noEmployee.contains(j)) {
                                        abstractModel.setFlag(4);
                                    }
                                    continue;
                                }
                                if (k == 1 || k == 2) {
                                    continue;
                                }
                                if (row.getCell(k) != null) {
                                try {
                                        double numericCellValue = row.getCell(k).getNumericCellValue();
                                        if (numericCellValue != 0.0) {
                                            map.put(String.valueOf(headMap.get(k)), numericCellValue);
                                        }
                                } catch (Exception e) {
                                    map.put(String.valueOf(headMap.get(k)), row.getCell(k).getStringCellValue());
                                }
                                }
                            }
                            // 保存记录
                            if (map.size() == 0) {
                                map.put("", "");
                            }
                            abstractModel.setResultMap(map);
                            if (StrUtil.isNotEmpty(abstractModel.getSubject())) {
                                    resultList.add(abstractModel);
                            }
                        }
                    }
                }
            }
        }catch (Exception e) {
            e.printStackTrace();
            System.out.println(e.getCause());
        }
        Map<Integer, List<AbstractModel>> collect = resultList.parallelStream().collect(Collectors.groupingBy(item -> item.getFlag()));

        return collect;
    }

    /**
     * 将处理好的数据转换输出
     * @param source
     */
    public static void transformSource(Map<Integer, List<AbstractModel>> source, int flag) {
        String filename = null;
        if (flag == 1) {
           filename = "公司园（幼儿）模板.xlsx";
        }
        if (flag == 2) {
            filename = "非公司园（幼儿）模板.xlsx";
        }
        if (flag == 3) {
            filename = "公司园（老师）模板.xlsx";
        }
        if (flag == 4) {
            filename = "非公司园（老师）模板.xlsx";
        }
        try (InputStream is= getFileInputStream(System.getProperty("user.dir") + "\\" + filename)) {
                try (Workbook workbook = XSSFWorkbookFactory.create(is)) {
                    List<AbstractModel> abstractModels = source.get(flag);
                    // 按周分组
                    Map<Integer, List<AbstractModel>> collect = abstractModels.parallelStream().collect(Collectors.groupingBy(item -> item.getWeek()));
                    for (int i = 0; i < workbook.getNumberOfSheets() - 1; i++) {
                        int weekOffset = 0;
                        //当前周的数据
                        List<AbstractModel> weekSource = collect.get(i);
                        Sheet sheetAt = workbook.getSheetAt(i);
                        if (i == 0) {
                            sheetAt.getRow(1).getCell(1).setCellValue(weekSource.get(0).getCurrentDate());
                        }
                        for (int j = 0; j < sheetAt.getLastRowNum(); j++) {
                            //
                            if (j == 4) {
                                Map<String, Object> resultMapOne = getWeekSourceForSpecial(weekSource, weekOffset);
                                String subjectOne = getSubjectForSpecial(weekSource, weekOffset);
                                if (StrUtil.isNotEmpty(subjectOne)) {
                                    setCellSubject(sheetAt, subjectOne, j, 0);
                                }
                                Map<String, Object> resultMapTwo = getWeekSourceForSpecial(weekSource, ++weekOffset);
                                String subjectTwo = getSubjectForSpecial(weekSource, weekOffset);
                                if (StrUtil.isNotEmpty(subjectTwo)) {
                                    setCellSubject(sheetAt, subjectTwo, j, 7);
                                }
                                setCellValue(resultMapOne, sheetAt, 4, 1);
                                setCellValue(resultMapTwo, sheetAt, 4, 8);
                            }
                            if (j == 54) {
                                Map<String, Object> resultMapOne = getWeekSourceForSpecial(weekSource, ++weekOffset);
                                String subjectOne = getSubjectForSpecial(weekSource, weekOffset);
                                if (StrUtil.isNotEmpty(subjectOne)) {
                                    setCellSubject(sheetAt, subjectOne, j, 0);
                                }
                                Map<String, Object> resultMapTwo = getWeekSourceForSpecial(weekSource, ++weekOffset);
                                String subjectTwo = getSubjectForSpecial(weekSource, weekOffset);
                                if (StrUtil.isNotEmpty(subjectTwo)) {
                                    setCellSubject(sheetAt, subjectTwo, j, 7);
                                }
                                setCellValue(resultMapOne, sheetAt, 54, 1);
                                setCellValue(resultMapTwo, sheetAt, 54, 8);
                            }
                            if (j == 107) {
                                Map<String, Object> resultMapOne = getWeekSourceForSpecial(weekSource, ++weekOffset);
                                String subjectOne = getSubjectForSpecial(weekSource, weekOffset);
                                if (StrUtil.isNotEmpty(subjectOne)) {
                                    setCellSubject(sheetAt, subjectOne, j, 0);
                                }
                                Map<String, Object> resultMapTwo = getWeekSourceForSpecial(weekSource, ++weekOffset);
                                String subjectTwo = getSubjectForSpecial(weekSource, weekOffset);
                                if (StrUtil.isNotEmpty(subjectTwo)) {
                                    setCellSubject(sheetAt, subjectTwo, j, 7);
                                }
                                setCellValue(resultMapOne, sheetAt, 107, 1);
                                setCellValue(resultMapTwo, sheetAt, 107, 8);
                            }
                            if (j == 151) {
                                Map<String, Object> resultMapOne = getWeekSourceForSpecial(weekSource, ++weekOffset);
                                String subjectOne = getSubjectForSpecial(weekSource, weekOffset);
                                if (StrUtil.isNotEmpty(subjectOne)) {
                                    setCellSubject(sheetAt, subjectOne, j, 0);
                                }
                                Map<String, Object> resultMapTwo = getWeekSourceForSpecial(weekSource, ++weekOffset);
                                String subjectTwo = getSubjectForSpecial(weekSource, weekOffset);
                                if (StrUtil.isNotEmpty(subjectTwo)) {
                                    setCellSubject(sheetAt, subjectTwo, j, 7);
                                }
                                setCellValue(resultMapOne, sheetAt, 151, 1);
                                setCellValue(resultMapTwo, sheetAt, 151, 8);
                            }
                            if (j == 195) {
                                Map<String, Object> resultMapOne = getWeekSourceForSpecial(weekSource, ++weekOffset);
                                String subjectOne = getSubjectForSpecial(weekSource, weekOffset);
                                if (StrUtil.isNotEmpty(subjectOne)) {
                                    setCellSubject(sheetAt, subjectOne, j, 0);
                                }
                                Map<String, Object> resultMapTwo = getWeekSourceForSpecial(weekSource, ++weekOffset);
                                String subjectTwo = getSubjectForSpecial(weekSource, weekOffset);
                                if (StrUtil.isNotEmpty(subjectTwo)) {
                                    setCellSubject(sheetAt, subjectTwo, j, 7);
                                }
                                setCellValue(resultMapOne, sheetAt, 195, 1);
                                setCellValue(resultMapTwo, sheetAt, 195, 8);
                            }
                            if (j == 239) {
                                Map<String, Object> resultMapOne = getWeekSourceForSpecial(weekSource, ++weekOffset);
                                String subjectOne = getSubjectForSpecial(weekSource, weekOffset);
                                if (StrUtil.isNotEmpty(subjectOne)) {
                                    setCellSubject(sheetAt, subjectOne, j, 0);
                                }
                                Map<String, Object> resultMapTwo = getWeekSourceForSpecial(weekSource, ++weekOffset);
                                String subjectTwo = getSubjectForSpecial(weekSource, weekOffset);
                                if (StrUtil.isNotEmpty(subjectTwo)) {
                                    setCellSubject(sheetAt, subjectTwo, j, 7);
                                }
                                setCellValue(resultMapOne, sheetAt, 239, 1);
                                setCellValue(resultMapTwo, sheetAt, 239, 8);
                            }
                        }
                    }
                    workbook.write(new FileOutputStream(new File(System.getProperty("user.dir")  + "\\newFile\\" + filename.replace("模板", ""))));
                }
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println(filename + "======> 生成出错 <======");
            System.out.println(e.getCause());
        }
        System.out.println(filename + "=====> 生成完毕 <======");
    }

    public static void main(String[] args) {
        Map<Integer, List<AbstractModel>> source = getSource();
        for (int i = 1; i < 5; i++) {
            transformSource(source, i);
        }
    }


    /**
     * 设置cell值
     * @param map
     * @param sheet
     * @param x
     * @param y
     */
    public static void setCellValue(Map<String, Object> map, Sheet sheet, int x, int y) {
        if (CollectionUtil.isEmpty(map)) {
            return;
        }
            for (Map.Entry<String, Object> entry: map.entrySet()
            ) {
                Cell cell = sheet.getRow(x).getCell(y);
                if (cell == null) {
                    return;
                }
                cell.setCellValue(entry.getKey());
                try {
                    sheet.getRow(x).getCell(y + 1).setCellValue((Double) entry.getValue());
                } catch (Exception e) {
                    sheet.getRow(x).getCell(y + 1).setCellValue((String) entry.getValue());
                }
                ++x;
            }
    }

    private static Map<String, Object> getWeekSourceForSpecial(List<AbstractModel> source, int index) {
        try {
            AbstractModel abstractModel = source.get(index);
            return abstractModel.getResultMap();
        } catch (Exception e) {
            return null;
        }
    }

    private static String getSubjectForSpecial(List<AbstractModel> source, int index) {
        try {
            AbstractModel abstractModel = source.get(index);
            return abstractModel.getSubject() + "送菜明细表";
        } catch (Exception e) {
            return null;
        }
    }

    private static InputStream getFileInputStream(String url) {
        File file = new File(url);
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return IoUtil.toBuffered(fileInputStream);
    }

    private static void setCellSubject(Sheet sheet, String subject, int j, int k) {
        Cell cell = sheet.getRow(j - 4).getCell(k);
        if (cell != null) {
            cell.setCellValue(subject);
        }
    }
}
