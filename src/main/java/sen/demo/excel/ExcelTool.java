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

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Locale;

/**
 * @author Huang Sen
 */
public class ExcelTool {

    public static void main(String[] args) throws Exception {
        String filePath = "D:\\t1.xlsx";
        //读取Excel中特定列的号码列表
        readExcel(filePath);

        //将号码添加到URL，根据URL远程调用后获取号码相关信息

        //解析相关信息中的归属地

        //将归属地存储至列表中

        //将归属地列表信息写入Excel
    }

    private static void readExcel(String filePath) throws Exception {
        File file = new File(filePath);
        Workbook workbook = WorkbookFactory.create(file);
        Sheet originalTable = workbook.getSheet("原表");
        //指定手机号码列数
        Iterator<Row> rowIterator = originalTable.rowIterator();
        int rowIndex = 0;
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (row == null||row.getCell(0).getStringCellValue().equals("部门")) {
                continue;
            }
            int phoneNumberColumnIndex = 3;
            int cityColumnIndex = 4;
            writeCity(row, phoneNumberColumnIndex, cityColumnIndex);
            writeCity(row, 5, 7);
            writeCity(row, 6, 8);
            System.out.println(rowIndex++);
        }
        System.out.println("开始写新的Excel");
        FileOutputStream fos = new FileOutputStream("D:\\t2.xls");
        workbook.write(fos);
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
