package com.example.exceldemo;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

@Controller
public class ExcelUploadController {

    @GetMapping("/excel")
    public String main() {
        return "uploadExcel";
    }

    @RequestMapping("/excel/read")
    public String readExcel(@RequestParam("file") MultipartFile file, Model model)
            throws IOException {
        // 파일종류 에러처리
        String extension = FilenameUtils.getExtension(file.getOriginalFilename());
        if (!extension.equals("xlsx") && !extension.equals("xls")) {
            throw new IOException("엑셀파일만 업로드 해주세요.");
        }

        Workbook workbook = null;
        if (extension.equals("xlsx")) {
            workbook = new XSSFWorkbook(file.getInputStream());
        } else if (extension.equals("xls")) {
            workbook = new HSSFWorkbook(file.getInputStream());
        }

        Sheet worksheet = workbook.getSheetAt(0);
        Row row;
        String[] strDate = new String[5]; // 날짜
        String[] strTime = {"아침", "점심", "저녁"}; // 시간
        String[] strCtgry = new String[10]; // 카테고리
        String[] strCtgryDessert = {"후식", "후식", "후식", "후식", "후식"};
        String[] strMenu = new String[10]; // 메뉴
        String[] strDessert = new String[5]; //후식

        List<ExcelData> dataList = new ArrayList<>();
        ExcelData data = null;

        /**
         * 아침 2행~9행
         * 2행 : 날짜
         * 3행 : 카테고리
         * 4~8행 : 메뉴
         * 9행 : 후식
         */
        for (int c = 1; c <= 10; c++) {
            for (int r = 2; r < 10; r++) {
                row = worksheet.getRow(r);
                if (row.getRowNum() == 2 && (c % 2) == 0) { //날짜
                    Date date = row.getCell(c).getDateCellValue();
                    strDate[(c / 2) - 1] = new SimpleDateFormat("yyyy-MM-dd").format(date);
                } else if (row.getRowNum() == 3) { //카테고리
                    strCtgry[c - 1] = row.getCell(c).getStringCellValue();
                } else if (row.getRowNum() == 9 && (c % 2) == 1) { //후식
                    strDessert[(c / 2)] = row.getCell(c).getStringCellValue();
                } else if (row.getRowNum() > 3 && row.getRowNum() < 9) { //메뉴
                    if (strMenu[c - 1] == null) {
                        strMenu[c - 1] = "" + row.getCell(c).getStringCellValue();
                    } else {
                        strMenu[c - 1] = strMenu[c - 1] + "\n" + row.getCell(c).getStringCellValue();
                    }
                }
            }
        }

        /**
         * 메뉴 + 후식 배열 합치기
         */
        List<String> list1 = new ArrayList(Arrays.asList(strMenu));
        List<String> list2 = new ArrayList(Arrays.asList(strDessert));
        list1.addAll(list2);
        String[] finalMenu = list1.toArray(new String[0]);

        List<String> list3 = new ArrayList(Arrays.asList(strCtgry));
        List<String> list4 = new ArrayList(Arrays.asList(strCtgryDessert));
        list3.addAll(list4);
        String[] finalCtgry = list3.toArray(new String[0]);

        dataList = insertDataList(strDate, dataList, finalMenu, finalCtgry, strTime, 0);

        /**
         * 점심 10행~18행
         * 10행 : 카테고리
         * 11~15행 : 메뉴
         * 16~17행 : 후식
         **/
        strCtgry = new String[10]; // 카테고리
        strMenu = new String[10]; // 메뉴
        strDessert = new String[5]; //후식

        for (int c = 1; c <= 10; c++) {
            for (int r = 10; r < 19; r++) {
                row = worksheet.getRow(r);
                if (row.getRowNum() == 10) { //카테고리
                    strCtgry[c - 1] = row.getCell(c).getStringCellValue();
                } else if (row.getRowNum() >= 17 && (c % 2) == 1) { //후식
                    if (strDessert[(c / 2)] == null) {
                        strDessert[(c / 2)] = "" + row.getCell(c).getStringCellValue();
                    } else {
                        strDessert[(c / 2)] = strDessert[(c / 2)] + "\n" + row.getCell(c).getStringCellValue();
                    }
                } else if (row.getRowNum() > 10 && row.getRowNum() < 16) { //메뉴
                    if (strMenu[c - 1] == null) {
                        strMenu[c - 1] = "" + row.getCell(c).getStringCellValue();
                    } else {
                        strMenu[c - 1] = strMenu[c - 1] + "\n" + row.getCell(c).getStringCellValue();
                    }
                }
            }
        }

        /**
         * 후식 배열 합치기
         */
        list1 = new ArrayList(Arrays.asList(strMenu));
        list2 = new ArrayList(Arrays.asList(strDessert));
        list1.addAll(list2);
        finalMenu = list1.toArray(new String[0]);

        list3 = new ArrayList(Arrays.asList(strCtgry));
        list4 = new ArrayList(Arrays.asList(strCtgryDessert));
        list3.addAll(list4);
        finalCtgry = list3.toArray(new String[0]);

        dataList = insertDataList(strDate, dataList, finalMenu, finalCtgry, strTime, 1);

        /**
         * 저녁 19행~27행
         * 19행 : 카테고리
         * 20~25행 : 메뉴
         * 26행 : 후식
         */
        strCtgry = new String[10]; // 카테고리
        strMenu = new String[10]; // 메뉴
        strDessert = new String[5]; //후식

        for (int c = 1; c <= 10; c++) {
            for (int r = 19; r < 28; r++) {
                row = worksheet.getRow(r);
                if (row.getRowNum() == 19) { //카테고리
                    strCtgry[c - 1] = row.getCell(c).getStringCellValue();
                } else if (row.getRowNum() >= 26 && (c % 2) == 1) { //후식
                    if (strDessert[(c / 2)] == null) {
                        strDessert[(c / 2)] = "" + row.getCell(c).getStringCellValue();
                    } else {
                        strDessert[(c / 2)] = strDessert[(c / 2)] + "\n" + row.getCell(c).getStringCellValue();
                    }
                } else if (row.getRowNum() > 19 && row.getRowNum() < 26) { //메뉴
                    if (strMenu[c - 1] == null) {
                        strMenu[c - 1] = "" + row.getCell(c).getStringCellValue();
                    } else {
                        strMenu[c - 1] = strMenu[c - 1] + "\n" + row.getCell(c).getStringCellValue();
                    }
                }
            }
        }

        /**
         * 후식 배열 합치기
         */
        list1 = new ArrayList(Arrays.asList(strMenu));
        list2 = new ArrayList(Arrays.asList(strDessert));
        list1.addAll(list2);
        finalMenu = list1.toArray(new String[0]);

        list3 = new ArrayList(Arrays.asList(strCtgry));
        list4 = new ArrayList(Arrays.asList(strCtgryDessert));
        list3.addAll(list4);
        finalCtgry = list3.toArray(new String[0]);

        dataList = insertDataList(strDate, dataList, finalMenu, finalCtgry, strTime, 2);

        // 불필요한 데이터 삭제
        dataList = dataList.stream().filter(
                excelData -> !(excelData.getMenu().replaceAll("\\n", "").equals("")
                || excelData.getMenu() == null
                || excelData.getCategory().contains("[가정의 날]")
                || excelData.getCategory().contains("[행복DAY]"))).collect(Collectors.toList());

        // 데이터 앞뒤 공백 제거
        dataList.stream().forEach(excelData -> excelData.setMenu(excelData.getMenu().trim()));

        model.addAttribute("datas", dataList);
        return "excelList";
    }

    private List<ExcelData> insertDataList(String[] strDate, List<ExcelData> dataList, String[] finalMenu, String[] finalCtgry, String[] strTime, int i2) {
        ExcelData data;
        for (int i = 0; i < 15; i++) {
            data = new ExcelData();
            if (i > 0 && i < 10) {
                data.setDate(strDate[i / 2]);
            } else if (i > 0 && i >= 10) {
                data.setDate(strDate[i - 10]);
            } else {
                data.setDate(strDate[i]);
            }

            data.setTime(strTime[i2]);
            data.setCategory(finalCtgry[i]);
            data.setMenu(finalMenu[i]);
            dataList.add(data);
        }
        return dataList;
    }
}
