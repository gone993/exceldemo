package com.example.exceldemo;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

@Controller
public class ExcelUploadController2 {

    @GetMapping("/excel2")
    public String main() {
        return "uploadExcel";
    }

    @RequestMapping("/excel/read2")
    public String readExcel(@RequestParam("file") MultipartFile file, Model model)
            throws IOException {

        List<ExcelData> dataList = new ArrayList<>();

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

        /**
         * 아침
         */
        for (int c = 1; c <= 10; c++) {
            for (int r = 2; r < 19; r++) {
                row = worksheet.getRow(r);
                if (row.getRowNum() == 2 && (c % 2) == 0) { //날짜
                    Date date = row.getCell(c).getDateCellValue();
                    strDate[(c / 2) - 1] = new SimpleDateFormat("yyyy-MM-dd").format(date);
                } else if (row.getRowNum() == 3 || row.getRowNum() == 10) { //카테고리
                    strCtgry[c - 1] = row.getCell(c).getStringCellValue();
                } else if ((row.getRowNum() == 9 && (c % 2) == 1) || (row.getRowNum() >= 17 && (c % 2) == 1)) { //후식
                    //strDessert[(c / 2)] = row.getCell(c).getStringCellValue();
                    if (strDessert[(c / 2)] == null) {
                        strDessert[(c / 2)] = "" + row.getCell(c).getStringCellValue();
                    } else {
                        strDessert[(c / 2)] = strDessert[(c / 2)] + "\n" + row.getCell(c).getStringCellValue();
                    }
                } else if ((row.getRowNum() > 3 && row.getRowNum() < 9)||(row.getRowNum() > 10 && row.getRowNum() < 16)) { //메뉴
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

        for (int i = 0; i < 15; i++) {
            ExcelData data = new ExcelData();

            if (i > 0 && i < 10) {
                data.setDate(strDate[i / 2]);
            } else if (i > 0 && i >= 10) {
                data.setDate(strDate[i - 10]);
            } else {
                data.setDate(strDate[i]);
            }

            data.setTime(strTime[0]);

            data.setCategory(finalCtgry[i]);
            data.setMenu(finalMenu[i]);

            dataList.add(data);
        }


        /**
         * 저녁
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

        for (int i = 0; i < 15; i++) {
            ExcelData data = new ExcelData();

            if (i > 0 && i < 10) {
                data.setDate(strDate[i / 2]);
            } else if (i > 0 && i >= 10) {
                data.setDate(strDate[i - 10]);
            } else {
                data.setDate(strDate[i]);
            }

            data.setTime(strTime[2]);

            data.setCategory(finalCtgry[i]);
            data.setMenu(finalMenu[i]);

            dataList.add(data);
        }

        model.addAttribute("datas", dataList);
        return "excelList";
    }
}
