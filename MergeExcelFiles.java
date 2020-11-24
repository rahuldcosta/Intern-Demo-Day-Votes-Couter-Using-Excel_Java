package demoDay;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class MergeExcelFiles {

    public static final String INTERN_DEMO_DAY_RESULTS_SHEET_NAME = "Intern Demo Day Results";
    public static final String TOTAL_VOTES = "Total Votes";
    public static final String AVERAGE_RATING = "Average Rating";
    public static final String INTERN_DEMO_DAY_RESULTS_FILE_NAME = "consolidatedTeamRatingResults.xlsx";
    public static final String EMPLOYEE_VOTING_WINNERS = "Employee Voting Winners";

    public static void main(String[] args) {
        Map<String, Integer> employeeVotes = new HashMap<>();
        Map<String, List<String>> timeStampMap = new HashMap<>();
        String exportedExcelPath = "C:\\exportedExcels\\";
        List<String> teamFileNames = new ArrayList<>();
        File[] files = new File(exportedExcelPath).listFiles();
        XSSFWorkbook resultsWorkbook = new XSSFWorkbook();
        cleanUpFinalResultFile(exportedExcelPath);
        getFilesInDirectory(teamFileNames, files);
        try {
            FileOutputStream out = new FileOutputStream(new File(exportedExcelPath + INTERN_DEMO_DAY_RESULTS_FILE_NAME));
            for (String excelFileName : teamFileNames
            ) {
                FileInputStream currentTeamExcelFile = null;
                currentTeamExcelFile = new FileInputStream(new File(exportedExcelPath + excelFileName));
                XSSFWorkbook currentExcelWorkbook = new XSSFWorkbook(currentTeamExcelFile);
                XSSFSheet currentFileSheet = currentExcelWorkbook.getSheetAt(0);
                XSSFSheet teamNameSheet = resultsWorkbook.createSheet(getTeamSheetName(excelFileName));
                Iterator<Row> rowIterator = currentFileSheet.iterator();
                int rowCounter = 2; // to skip first two rows as they are used for total votes and rating along with a blank row
                int colCounter = 0;
                XSSFRow outputFileRow;
                XSSFCell outputFileCell;
                while (rowIterator.hasNext()) {
                    Row currentExcelRow = rowIterator.next();
                    outputFileRow = teamNameSheet.createRow(rowCounter);
                    colCounter = 0;
                    outputFileCell = outputFileRow.createCell(0);
                    String userEmailId = currentExcelRow.getCell(3).getStringCellValue();
                    outputFileCell.setCellValue(userEmailId);
                    outputFileCell = outputFileRow.createCell(1);
                    outputFileCell.setCellValue(currentExcelRow.getCell(4).getStringCellValue());
                    outputFileCell = outputFileRow.createCell(2);
                    if (rowCounter > 2) {
                        outputFileCell.setCellValue(currentExcelRow.getCell(5).getNumericCellValue());

                        if (employeeVotes.containsKey(userEmailId)) {
                            employeeVotes.put(userEmailId, employeeVotes.get(userEmailId) + 1);
                            timeStampMap.get(userEmailId).add("start_" + currentExcelRow.getCell(1).getDateCellValue());
                            timeStampMap.get(userEmailId).add("end_" + currentExcelRow.getCell(2).getDateCellValue());
                        } else {
                            employeeVotes.put(userEmailId, 1);
                            timeStampMap.put(userEmailId, new ArrayList<>());
                            timeStampMap.get(userEmailId).add("start_" + currentExcelRow.getCell(1).getDateCellValue());
                            timeStampMap.get(userEmailId).add("end_" + currentExcelRow.getCell(2).getDateCellValue());
                        }
                    } else
                        outputFileCell.setCellValue("Rating");
                    rowCounter++;
                }
                outputFileRow = teamNameSheet.createRow(0);
                outputFileCell = outputFileRow.createCell(0);
                outputFileCell.setCellValue(TOTAL_VOTES);
                outputFileCell = outputFileRow.createCell(1);
                outputFileCell.setCellValue(currentFileSheet.getPhysicalNumberOfRows() - 1);
                outputFileCell = outputFileRow.createCell(2);
                outputFileCell.setCellValue(AVERAGE_RATING);
                outputFileCell = outputFileRow.createCell(3);
                outputFileCell.setCellFormula("ROUND(AVERAGE(C4:C" + (currentFileSheet.getPhysicalNumberOfRows() + 2) + "),3)");
                outputFileRow = teamNameSheet.createRow(1);

            }
            populateResultsSheetData(resultsWorkbook, teamFileNames);
            populateEmployeeVoteCountResult(resultsWorkbook, employeeVotes, timeStampMap);
            resultsWorkbook.setSheetOrder(INTERN_DEMO_DAY_RESULTS_SHEET_NAME, 0);
            resultsWorkbook.setSheetOrder(EMPLOYEE_VOTING_WINNERS, 1);
            resultsWorkbook.setSelectedTab(0);
            resultsWorkbook.setActiveSheet(0);
            resultsWorkbook.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static void populateEmployeeVoteCountResult(XSSFWorkbook resultsWorkbook, Map<String, Integer> employeeVotes, Map<String, List<String>> timeStampMap) {
        XSSFRow outputFileRow;
        XSSFCell outputFileCell;
        XSSFSheet sheet = resultsWorkbook.createSheet(EMPLOYEE_VOTING_WINNERS);
        int rowCount = 1;
        for (Map.Entry<String, Integer> entry : employeeVotes.entrySet()) {
            outputFileRow = sheet.createRow(rowCount);
            outputFileCell = outputFileRow.createCell(0);
            outputFileCell.setCellValue(entry.getKey());
            outputFileCell = outputFileRow.createCell(1);
            outputFileCell.setCellValue(entry.getValue());

            List<Date> dateTimeList = new ArrayList<>();
            SimpleDateFormat formatter = new SimpleDateFormat("EE MMM dd HH:mm:ss z yyyy",
                    Locale.ENGLISH);

            for (Map.Entry<String, List<String>> timeStampEntry : timeStampMap.entrySet()) {

                if (timeStampEntry.getKey().equals(entry.getKey())) {

                    for (String date :
                            timeStampEntry.getValue()) {
                        try {
                            Date currentDate = formatter.parse(date.replaceAll("start_", "").replaceAll("end_", ""));
                            dateTimeList.add(currentDate);
                        } catch (ParseException e) {
                            e.printStackTrace();
                        }
                    }
                    outputFileCell = outputFileRow.createCell(2);
                    outputFileCell.setCellValue(findTotalVotingTimeSpendByEmployee(Collections.max(dateTimeList), Collections.min(dateTimeList)));
                    break;
                }
            }


            rowCount++;
        }


    }

    private static String findTotalVotingTimeSpendByEmployee(Date max, Date min) {
        long diff = max.getTime() - min.getTime();

        long diffSeconds = diff / 1000 % 60;
        long diffMinutes = diff / (60 * 1000) % 60;
        long diffHours = diff / (60 * 60 * 1000) % 24;

        return diffHours + " hours, " + diffMinutes + " minutes, " + diffSeconds + " seconds.";
    }

    private static void populateResultsSheetData(XSSFWorkbook resultsWorkbook, List<String> fileNameList) {
        XSSFRow outputFileRow;
        XSSFCell outputFileCell;
        XSSFSheet sheet = resultsWorkbook.createSheet(INTERN_DEMO_DAY_RESULTS_SHEET_NAME);
        populateHeadersForResultsSheet(sheet);
        int rowCount = 1;
        for (String fileName : fileNameList
        ) {
            outputFileRow = sheet.createRow(rowCount);
            String teamName = getTeamSheetName(fileName);
            outputFileCell = outputFileRow.createCell(0);
            outputFileCell.setCellValue(teamName);
            outputFileCell = outputFileRow.createCell(1);
            outputFileCell.setCellFormula("'" + teamName + "'!B1");
            outputFileCell = outputFileRow.createCell(2);
            outputFileCell.setCellFormula("'" + teamName + "'!D1");
            rowCount++;
        }

    }

    private static void populateHeadersForResultsSheet(XSSFSheet sheet) {
        XSSFRow outputFileRow;
        XSSFCell outputFileCell;
        outputFileRow = sheet.createRow(0);
        outputFileCell = outputFileRow.createCell(0);
        outputFileCell.setCellValue("Team Name");
        outputFileCell = outputFileRow.createCell(1);
        outputFileCell.setCellValue(TOTAL_VOTES);
        outputFileCell = outputFileRow.createCell(2);
        outputFileCell.setCellValue(AVERAGE_RATING);
    }

    private static String getTeamSheetName(String excelFileName) {
        return excelFileName.split("\\.")[0].split("\\(")[0].replaceAll("_-", "").replaceFirst("Team", "");
    }

    private static void cleanUpFinalResultFile(String exportedExcelPath) {
        File outputFile = new File(exportedExcelPath + INTERN_DEMO_DAY_RESULTS_FILE_NAME);
        outputFile.delete();
    }

    private static void getFilesInDirectory(List<String> teamFileNames, File[] files) {
        for (File file : files) {
            if (file.isFile()) {
                teamFileNames.add(file.getName());
            }
        }
    }


}
