import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import java.io.*;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
    public static final String USER_AGENT = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.101 Safari/537.36";
    private static String searchTerm;
    private static int num;

    /**
     * Method to convert the {@link InputStream} to {@link String}
     *
     * @param is the {@link InputStream} object
     * @return the {@link String} object returned
     */
    public static String getString(InputStream is) {
        StringBuilder sb = new StringBuilder();

        BufferedReader br = new BufferedReader(new InputStreamReader(is));
        String line;
        try {
            while ((line = br.readLine()) != null) {
                sb.append(line);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            /** finally block to close the {@link BufferedReader} */
            if (br != null) {
                try {
                    br.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return sb.toString();
    }

    /**
     * The method will return the search page result in a {@link String} object
     *
     * @param path the google search query
     * @return the content as {@link String} object
     * @throws Exception
     */
    public static String getSearchContent(String path) throws Exception {
        final String agent = "Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)";
        URL url = new URL(path);
        final URLConnection connection = url.openConnection();
        /**
         * User-Agent is mandatory otherwise Google will return HTTP response
         * code: 403
         */
        connection.setRequestProperty("User-Agent", agent);
        final InputStream stream = connection.getInputStream();
        return getString(stream);
    }

    /**
     * Parse all links
     *
     * @param html the page
     * @return the list with all URLSs
     * @throws Exception
     */
    public static List<String> parseLinks(final String html) throws Exception {
        List<String> result = new ArrayList<String>();
        //      String pattern1 = "<h3 class=\"r\"><a href=\"/url?q=";
        String pattern1 = "<h3 class=\"r\"><a href=\"/url?q=";
        String pattern2 = "\">";
        Pattern p = Pattern.compile(Pattern.quote(pattern1) + "(.*?)" + Pattern.quote(pattern2));
        Matcher m = p.matcher(html);
        //<h3 class="r"><a href="https://www.aparat.com/v/FUbSf/%D8%A8%D9%86_%D8%AA%D9%86"
        //<h3 class="r"><a href="https://www.cartoonnetwork.com/games/ben-10/index.html"

        while (m.find()) {
            String domainName = m.group(0).trim();
            if (!domainName.contains("webcache.googleusercontent")) {

                /** remove the unwanted text */
                domainName = domainName.substring(domainName.indexOf("/url?q=") + 7);
                domainName = domainName.substring(0, domainName.indexOf("&amp;"));

                result.add(domainName);
            }
        }
        return result;
    }

    public static String getResultSiteTitle(String url) {
        try {
            URL url1 = new URL(url);
            return url1.getHost();
        } catch (MalformedURLException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * Does all the parsing job automatic using JSoup library
     * returns the result in a string
     * @param searchTerm keyword for searching
     * @param num number of desired results
     * @return Result
     */
    public static String getCompleteResult(String searchTerm , int num){
        setSearchTerm(searchTerm);
        setNum(num);
        int startIndex = 0;
        int rank = 1;
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("Google Search Parser\n");
        stringBuilder.append("Searching for: " + searchTerm+"\n\n");
        searchTerm = searchTerm.replaceAll(" ", "+");
        stringBuilder.append("Results:\n");
        while (rank < num) {
            String query = "https://www.google.com/search?q=" + searchTerm + "&num=50&start=" + startIndex;
            //Fetch the page
            Document doc = null;
            try {
                doc = Jsoup.connect(query).userAgent(USER_AGENT).get();
            } catch (IOException e) {
                e.printStackTrace();
            }

            //Traverse the results
            for (Element result : doc.select("h3.r a")) {
                if(rank > num)
                    break;

                final String title = result.text();
                final String url = result.attr("href");

                //Now do something with the results
                stringBuilder.append(String.format("%4d    %-150s    %-16s" , rank++ , url , getResultSiteTitle(url)));
                stringBuilder.append("\n");
            }
            startIndex += 50;
        }
        return stringBuilder.toString();
    }

    /**
     * Writes all the result in a file at classResoruse/Results/[currentTimeAndDate].txt
     * @param res String for results
     */
    public static void getResultInTextFile(String res){
        String timeStamp = new SimpleDateFormat("yyyy-MM-dd_HH mm ss").format(Calendar.getInstance().getTime());
        BufferedWriter bufferedWriter = null;
        try {
            bufferedWriter = new BufferedWriter(new FileWriter("Results\\"+timeStamp+".txt"));
            bufferedWriter.write(res);
            bufferedWriter.flush();
            bufferedWriter.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void getResultInXLSXFile(String result) throws IOException {
        /*
        String timeStamp = new SimpleDateFormat("yyyy-MM-dd_HH mm ss").format(Calendar.getInstance().getTime());
        String excelFileName = "Results\\XLSX\\" + timeStamp + ".xlsx";//name of excel file

        String sheetName = searchTerm;//name of sheet

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName) ;
        BufferedReader bufferedReader = new BufferedReader(new StringReader(result));

        //Read 4 lines to get to the main results
        for(int i = 0 ; i < 4 ; i++)
            bufferedReader.readLine();
        //iterating r number of rows
        for (int r=0;r < num; r++ )
        {
            XSSFRow row = sheet.createRow(r);
            String[] strings = bufferedReader.readLine().split("\\s+");

            //iterating c number of columns
            for (int c=0 ; c < 3; c++ )
            {
                XSSFCell cell = row.createCell(c);
                cell.setCellValue(strings[c+1]);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        //write this workbook to an Outputstream.
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
        */

        String[] columns = {"Rank", "URL", "Domain"};

        // Create a Workbook
        Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet(searchTerm);

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        int rowNum = 1;
        BufferedReader bufferedReader = new BufferedReader(new StringReader(result));

        //Read 4 lines to get to the main results
        for(int i = 0 ; i < 3 ; i++)
            bufferedReader.readLine();

        String line = null;
        while((line = bufferedReader.readLine()) != null) {
            Row row = sheet.createRow(rowNum++);

            String[] strings = bufferedReader.readLine().split("\\s+");

            row.createCell(0).setCellValue(strings[1]);

            row.createCell(1).setCellValue(strings[2]);

            row.createCell(2).setCellValue(strings[3]);
        }

        // Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        String timeStamp = new SimpleDateFormat("yyyy-MM-dd_HH mm ss").format(Calendar.getInstance().getTime());
        String excelFileName = "Results\\XLSX\\" + timeStamp + ".xlsx";//name of excel file
        FileOutputStream fileOut = new FileOutputStream(excelFileName);
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }

    /**
     * Change the searchTerm and num (Number of results) in the getCompleteResult method to get the desired results
     * use getResultInTextFile to get a .txt for results with current date and time
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        String sampleResult = getCompleteResult("پارس تامین",25);
        System.out.println(sampleResult);
        getResultInTextFile(sampleResult);
        getResultInXLSXFile(sampleResult);
    }

    public static String getSearchTerm() {
        return searchTerm;
    }

    public static void setSearchTerm(String searchTerm) {
        Main.searchTerm = searchTerm;
    }

    public static int getNum() {
        return num;
    }

    public static void setNum(int num) {
        Main.num = num;
    }
}
