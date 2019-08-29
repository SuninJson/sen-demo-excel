package sen.demo.excel;

import com.google.i18n.phonenumbers.NumberParseException;
import com.google.i18n.phonenumbers.PhoneNumberUtil;
import com.google.i18n.phonenumbers.Phonenumber;
import com.google.i18n.phonenumbers.geocoding.PhoneNumberOfflineGeocoder;
import org.apache.http.HttpResponse;
import org.apache.http.HttpStatus;
import org.apache.http.client.ClientProtocolException;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.apache.poi.ss.usermodel.*;
import sen.common.util.file.FileUtil;
import sen.common.util.regular.RegularUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

/**
 * @author Huang Sen
 */
public class ExcelPhoneNumberHandler {

    public static String[] specialCities = {"江苏", "浙江", "上海"};

    public static void main(String[] args) throws Exception {
        String directoryPath = "D:\\hx-temp";
        List<String> filePathList = FileUtil.getAllFileAbsolutePath(directoryPath, false);
        //读取Excel中特定列的号码列表
        for (String filePath : filePathList) {
            readExcel(filePath);
        }

        //将号码添加到URL，根据URL远程调用后获取号码相关信息

        //解析相关信息中的归属地

        //将归属地存储至列表中

        //将归属地列表信息写入Excel
    }

    private static void readExcel(String filePath) throws Exception {
        File file = new File(filePath);
        Workbook workbook = WorkbookFactory.create(file);
        Workbook newWorkbook = WorkbookFactory.create(file);
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            if (isEmptySheet(sheet)) {
                continue;
            }
            String newSheetName = "补充手机号归属地后的" + sheet.getSheetName();
            String specialCitySheetName = "处理后只包含江苏、浙江和上海的" + sheet.getSheetName();
            Sheet newSheet = newWorkbook.createSheet(newSheetName);
            Sheet specialCitySheet = newWorkbook.createSheet(specialCitySheetName);

            Iterator<Row> rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                handleNewSheet(newSheet, row);
            }
            for (Row newRow : newSheet) {
                if (rowContainSpecialCity(newRow)) {
                    sheetCopyRow(specialCitySheet, newRow);
                }
            }
        }
        System.out.println("开始写新的Excel：" + filePath);
        FileOutputStream fos = new FileOutputStream("D:\\hx\\" + file.getName() + ".xlsx");
        newWorkbook.write(fos);
        fos.close();
    }

    private static void sheetCopyRow(Sheet specialCitySheet, Row sourceRow) {
        Row row = specialCitySheet.createRow(specialCitySheet.getLastRowNum() + 1);
        for (Cell cell : sourceRow) {
            Cell newCell = row.createCell(cell.getColumnIndex());
            cellSetValue(newCell, getCellValue(cell));
        }
    }

    private static boolean rowContainSpecialCity(Row row) {
        for (Cell newCell : row) {
            if (isSpecialCity(String.valueOf(getCellValue(newCell)))) {
                return true;
            }
        }
        return false;
    }

    public static boolean isEmptySheet(Sheet sheet) {
        return sheet.getLastRowNum() == 0 && sheet.getPhysicalNumberOfRows() == 0;
    }

    private static void handleNewSheet(Sheet newSheet, Row row) {
        Row newRow = newSheet.createRow(row.getRowNum());
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            Cell newCell = newRow.createCell(cell.getColumnIndex());
            Object cellValueObj = getCellValue(cell);
            cellSetValue(newCell, cellValueObj);
            if (cellValueObj instanceof String) {
                String cellValue = (String) cellValueObj;
                if (RegularUtil.isMobile(cellValue) || RegularUtil.isPhone(cellValue)) {
                    Cell lastCell = newRow.createCell(newRow.getLastCellNum() + 6);
                    String city = getCity(cellValue);
                    lastCell.setCellValue(city);
                }
            }
        }
    }

    private static void cellSetValue(Cell cell, Object value) {
        if (null == value) {
            //为空不处理
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof Double) {
            cell.setCellValue(Double.valueOf(String.valueOf(value)));
        } else if (value instanceof BigDecimal) {
            cell.setCellValue(((BigDecimal) value).toPlainString());
        } else {
            System.out.println("未处理类型：" + value.getClass());
        }
    }

    private static boolean isSpecialCity(String city) {
        for (String specialCity : specialCities) {
            if (city.contains(specialCity)) {
                return true;
            }
        }
        return false;
    }

    @SuppressWarnings("deprecation")
    public static Object getCellValue(Cell cell) {
        try {
            if (cell == null) {
                return "";
            }
            Object obj = null;
            switch (cell.getCellTypeEnum()) {
                case BOOLEAN:
                    obj = cell.getBooleanCellValue();
                    break;
                case ERROR:
                    obj = cell.getErrorCellValue();
                    break;
                case FORMULA:
                    try {
                        obj = String.valueOf(cell.getStringCellValue());
                    } catch (IllegalStateException e) {
                        String valueOf = String.valueOf(cell.getNumericCellValue());
                        BigDecimal bd = new BigDecimal(Double.valueOf(valueOf));
                        bd = bd.setScale(2, RoundingMode.HALF_UP);
                        obj = bd;
                    }
                    break;
                case NUMERIC:
                    obj = new BigDecimal(cell.getNumericCellValue()).toPlainString();
                    break;
                case STRING:
                    String value = String.valueOf(cell.getStringCellValue());
                    value = value.replace(" ", "");
                    value = value.replace("\n", "");
                    value = value.replace("\t", "");
                    obj = value;
                    break;
                default:
                    break;
            }
            return obj;
        } catch (Exception e) {
            return "错误数据";
        }
    }


    private static void writeCity(Row row, int phoneNumberColumnIndex, int cityColumnIndex) {
        Cell phoneNumberCell = row.getCell(phoneNumberColumnIndex);
        String phoneNumber;
        try {
            phoneNumber = phoneNumberCell.getStringCellValue();
        } catch (IllegalStateException e) {
            phoneNumber = String.valueOf(phoneNumberCell.getNumericCellValue());
        }
        String city = getCity(phoneNumber);
        Cell cityCell = row.getCell(cityColumnIndex);
        cityCell.setCellValue(city);
    }

    public static String getCity(String phoneNum) {
        PhoneNumberUtil phoneUtil = PhoneNumberUtil.getInstance();
        PhoneNumberOfflineGeocoder phoneNumberOfflineGeocoder = PhoneNumberOfflineGeocoder.getInstance();
        String language = "CN";
        Phonenumber.PhoneNumber referencePhonenumber = null;
        try {
            referencePhonenumber = phoneUtil.parse(phoneNum, language);
        } catch (NumberParseException e) {
            return "";
        }
        //手机号码归属城市 city
        return phoneNumberOfflineGeocoder.getDescriptionForNumber(referencePhonenumber, Locale.CHINA);
    }

    /**
     * 从www.ip138.com返回的结果网页内容中获取手机号码归属地,结果为：省份 城市
     * 选用这个的原因。是这个在尝试的数个获取api里面是比较精确的。
     *
     * @return
     */
    public static String getCityUrl(String mobile) {
        String url = "http://www.ip138.com:8080/search.asp";
        StringBuffer sb = new StringBuffer(url);
        sb.append("?mobile=" + mobile);
        sb.append("&action=mobile");
        // 指定get请求
        HttpGet httpGet = new HttpGet(sb.toString());
        // 创建httpclient
        CloseableHttpClient httpClient = HttpClients.createDefault();
        // 发送请求
        HttpResponse httpResponse;
        //返回的json
        String result = null;
        try {
            httpResponse = httpClient.execute(httpGet);
            // 验证请求是否成功
            if (httpResponse.getStatusLine().getStatusCode() == HttpStatus.SC_OK) {
                // 得到请求响应信息
                String str = EntityUtils.toString(httpResponse.getEntity(),
                        "GB2312");
                // 返回json
                if (str != null && !str.equals("")) {
                    result = parseMobileFrom(str);
                }
            }
        } catch (ClientProtocolException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return result;
    }


    public static String parseMobileFrom(String htmlSource) {
        String result = "";
        String[] htmls = htmlSource.split("\n");
        for (int i = 0; i < htmls.length; i++) {
            String thisHtml = htmls[i];
            if (thisHtml.indexOf("卡号归属地") > 0) {
                if (thisHtml.indexOf("tdc2") > 0) {
                    thisHtml = thisHtml.substring(0, thisHtml.lastIndexOf("<"));
                    result = thisHtml.substring(thisHtml.lastIndexOf(">") + 1);
                } else {
                    thisHtml = htmls[i + 1];
                    thisHtml = thisHtml.substring(0, thisHtml.lastIndexOf("<"));
                    result = thisHtml.substring(thisHtml.lastIndexOf(">") + 1);
                }
            }
        }
        return result.replaceAll("&nbsp;", "");
    }

}
