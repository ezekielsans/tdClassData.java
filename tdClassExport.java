/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package bean;

import com.liferay.faces.portal.context.LiferayFacesContext;
import com.liferay.portal.theme.ThemeDisplay;
import java.io.File;
import java.io.FileOutputStream;
import java.io.Serializable;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import javax.faces.application.FacesMessage;
import javax.faces.bean.ManagedBean;
import javax.faces.bean.ManagedProperty;
import javax.faces.bean.SessionScoped;
import javax.faces.context.FacesContext;
//import model.ClassTdSummary;
//import class_summary.ClassSummary;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Font;

/**
 *
 * @author misteam
 */
@ManagedBean
@SessionScoped
public class ClassTdData implements Serializable {

    /**
     * Creates a new instance of ClassTdData
     */
    public ClassTdData() {
    }

    /*
     * properties
     */
    @ManagedProperty(value = "#{customEntityManagerFactory}")
    private CustomEntityManagerFactory customEntityManagerFactory;
    @ManagedProperty(value = "#{accountsWithSubsidiaryData}")
    private AccountsWithSubsidiaryData accountsWithSubsidiaryData;
    @ManagedProperty(value = "#{printController}")
    private PrintController printController;
    @ManagedProperty(value = "#{customDate}")
    private CustomDate customDate;
    @ManagedProperty(value = "#{exportData}")
    private ExportData exportData;
    @ManagedProperty(value = "#{dataConvert}")
    private DataConvert dataConvert;
    private List<List<Object[]>> classTdSummary;
    private List<Object[]> classTdSummaryPrint;

    /*
     * getter setter
     */
    public CustomEntityManagerFactory getCustomEntityManagerFactory() {
        return customEntityManagerFactory == null ? customEntityManagerFactory = new CustomEntityManagerFactory() : customEntityManagerFactory;
    }

    public void setCustomEntityManagerFactory(CustomEntityManagerFactory customEntityManagerFactory) {
        this.customEntityManagerFactory = customEntityManagerFactory;
    }

    public AccountsWithSubsidiaryData getAccountsWithSubsidiaryData() {
        return accountsWithSubsidiaryData == null ? accountsWithSubsidiaryData = new AccountsWithSubsidiaryData() : accountsWithSubsidiaryData;
    }

    public void setAccountsWithSubsidiaryData(AccountsWithSubsidiaryData accountsWithSubsidiaryData) {
        this.accountsWithSubsidiaryData = accountsWithSubsidiaryData;
    }

    public PrintController getPrintController() {
        return printController == null ? printController = new PrintController() : printController;
    }

    public void setPrintController(PrintController printController) {
        this.printController = printController;
    }

    public CustomDate getCustomDate() {
        return customDate == null ? customDate = new CustomDate() : customDate;
    }

    public void setCustomDate(CustomDate customDate) {
        this.customDate = customDate;
    }

    public ExportData getExportData() {
        return exportData == null ? exportData = new ExportData() : exportData;
    }

    public void setExportData(ExportData exportData) {
        this.exportData = exportData;
    }

    public DataConvert getDataConvert() {
        return dataConvert == null ? dataConvert = new DataConvert() : dataConvert;
    }

    public void setDataConvert(DataConvert dataConvert) {
        this.dataConvert = dataConvert;
    }

    public List<List<Object[]>> getClassTdSummary() {
        return classTdSummary == null ? classTdSummary = new ArrayList<>() : classTdSummary;
    }

    public void setClassTdSummary(List<List<Object[]>> classTdSummary) {
        this.classTdSummary = classTdSummary;
    }

    public List<Object[]> getClassTdSummaryPrint() {
        return classTdSummaryPrint == null ? classTdSummaryPrint = new ArrayList<>() : classTdSummaryPrint;
    }

    public void setClassTdSummaryPrint(List<Object[]> classTdSummaryPrint) {
        this.classTdSummaryPrint = classTdSummaryPrint;
    }

    /*
     * methods
     */
    public void init() {
        String query;

        if (FacesContext.getCurrentInstance().isPostback() == false) {
            beanclear();
            getAccountsWithSubsidiaryData().beanclear();
            getAccountsWithSubsidiaryData().setShowSpecificColumn(Boolean.TRUE);
        }
//        hard-coded account_code
        query = "SELECT coa.accountCode FROM CoopFinChartOfAccounts coa WHERE coa.accountCode LIKE :accountCode AND coa.level = 5 ORDER BY coa.accountCode";
        getAccountsWithSubsidiaryData().setAccountCodes(getCustomEntityManagerFactory().getFinancialDbEntityManagerFactory().createEntityManager().createQuery(query).setParameter("accountCode", "211106%").getResultList());

        getAccountsWithSubsidiaryData().setAccountCodesInclude(null);

        for (int i = 0; i < getAccountsWithSubsidiaryData().getAccountCodes().size(); i++) {
            getAccountsWithSubsidiaryData().getAccountCodesInclude().add(i, false);
        }
    }

    public void beanclear() {
        FacesContext.getCurrentInstance().getExternalContext().getSessionMap().put("classTdData", null);
    }

    public String queryGenerator(String accountCodeParam, String selectedBatchParam) {
        String query;

        query = "SELECT x.row_number,  " //0
                + "x.slNo,  " //1
                + "x.account_no, "//2
                + "x.accountCode,  " //3
                + "x.acctName,  " //4
                + "x.balance, " //5
                + "x.acctStatus, "//6
                + "x.acctCreateDate, "//7
                + "x.postDate, "//8
                + "x.ctdNo,  " //9
                + "x.primary_holder, " //10
                + "x.sc_acctno "//11
                + "FROM ("
                + "SELECT row_number() OVER () AS row_number, "
                + "cls.slNo, "
                + "cls.account_no, "
                + "cls.accountCode, "
                + "cls.acctName, "
                + "cls.balance, "
                + "cls.acctStatus, "
                + "cls.acctCreateDate, "
                + "cls.postDate, "
                + "cls.ctdNo, "
                + "cls.primary_holder, "
                + "cls.acctno as sc_acctno "
                //inner sub
                + "FROM ("
                + "SELECT DISTINCT ON (apd.acctno) sl.td_sl_no AS slNo, "
                + "apd.acctno AS account_no, "
                + "apd.account_code AS accountCode, "
                + "cls.acct_name AS acctName, "
                + "sl.acct_balance AS balance, "
                + "apd.acct_status AS acctStatus, "
                + "apd.acct_create_date AS acctCreateDate, "
                + "sl.post_date AS postDate, "
                + "cls.ctd_no AS ctdNo, "
                + "apd.primary_holder , "
                + "b.acctno "
                + "FROM coop_fin_acct_profile_dtl apd "
                //moved left join for sc_acctno
                + "LEFT JOIN ( "
                + "SELECT a.primary_holder, "
                + "a.acctno "
                + "FROM coop_fin_acct_profile_dtl a "
                + "WHERE acctno ILIKE '%SC%'"
                + ") b ON apd.primary_holder = b.primary_holder  "
                + "JOIN coop_fin_class_td cls ON apd.acctno::text = cls.acctno::text "
                + "JOIN coop_fin_td_sl sl ON cls.acctno::text = sl.acctno::text ";
        if (getAccountsWithSubsidiaryData().getAcctCreateDateFrom() != null && getAccountsWithSubsidiaryData().getAcctCreateDateTo() != null) {
            query += "WHERE apd.acct_create_date BETWEEN '" + getCustomDate().formatDate(getAccountsWithSubsidiaryData().getAcctCreateDateFrom(), "yyyy-MM-dd") + "' AND '" + getCustomDate().formatDate(getAccountsWithSubsidiaryData().getAcctCreateDateTo(), "yyyy-MM-dd") + "' ";
        } else if (getAccountsWithSubsidiaryData().getAcctCreateDateFrom() != null && getAccountsWithSubsidiaryData().getAcctCreateDateTo() == null) {
            query += "WHERE apd.acct_create_date BETWEEN '" + getCustomDate().formatDate(getAccountsWithSubsidiaryData().getAcctCreateDateFrom(), "yyyy-MM-dd") + "' AND '" + getCustomDate().formatDate(getCustomDate().getCurrentDate(), "yyyy-MM-dd") + "' ";
        } else if (getAccountsWithSubsidiaryData().getAcctCreateDateFrom() == null && getAccountsWithSubsidiaryData().getAcctCreateDateTo() == null) {
            query += "WHERE apd.acct_create_date BETWEEN '1900-01-01' AND '" + getCustomDate().formatDate(getCustomDate().getCurrentDate(), "yyyy-MM-dd") + "' ";
        }

        if (getAccountsWithSubsidiaryData().getReportDate() == null) {
            query += "AND sl.post_date <= '" + getCustomDate().formatDate(getCustomDate().getCurrentDate(), "yyyy-MM-dd") + "' ";
        } else {
            query += "AND sl.post_date <= '" + getCustomDate().formatDate(getAccountsWithSubsidiaryData().getReportDate(), "yyyy-MM-dd") + "' ";
        }

        query += "AND apd.account_code = '" + accountCodeParam + "' ";

        if (getAccountsWithSubsidiaryData().getAmountFilter() != null && getAccountsWithSubsidiaryData().getAmountFilter().compareTo(BigDecimal.ZERO) == 1) {
            query += "AND sl.acct_balance <= " + getAccountsWithSubsidiaryData().getAmountFilter();
        }

//        new code below
//        query += "AND (apd.acct_status <> 'C') "
//                + "OR (apd.acct_status = 'C' "
//                + "AND sl.acct_balance <> 0) ";
//        new code above   
        query += " ORDER BY apd.acctno, sl.post_date DESC, sl.td_sl_no DESC) cls "
                + "WHERE (cls.acctStatus = 'A') "
                + "OR (cls.acctStatus = 'C' "
                + "AND cls.balance <> 0) "
                + "OR (cls.acctStatus = 'O' "
                + "AND cls.balance <> 0) "
                + "ORDER BY cls.acctno) x ";

        if (!selectedBatchParam.equals("ALL")) {
            query += "WHERE x. row_number " + selectedBatchParam;
        }

        return query;
    }

    public void runReport() {

        getAccountsWithSubsidiaryData().setSelectedAccountCode("");
        getAccountsWithSubsidiaryData().setAccountCodesFiltered(null);
        getAccountsWithSubsidiaryData().setNoOfRows(null);
        getAccountsWithSubsidiaryData().setGrandTotal(BigDecimal.ZERO);

        if (getAccountsWithSubsidiaryData().getAccountCodesInclude().contains(true)) {
            for (int i = 0; i < getAccountsWithSubsidiaryData().getAccountCodesInclude().size(); i++) {
                if (getAccountsWithSubsidiaryData().getAccountCodesInclude().get(i)) {
                    getAccountsWithSubsidiaryData().getAccountCodesFiltered().add(getAccountsWithSubsidiaryData().getAccountCodes().get(i));
                }
            }
        } else {
            getAccountsWithSubsidiaryData().setAccountCodesFiltered(getAccountsWithSubsidiaryData().getAccountCodes());
        }

        computeTotal();

        for (int i = 0; i < getAccountsWithSubsidiaryData().getAccountCodesFiltered().size(); i++) {
            getClassTdSummary().add(i, getCustomEntityManagerFactory().getFinancialDbEntityManagerFactory().createEntityManager().createNativeQuery(queryGenerator(getAccountsWithSubsidiaryData().getAccountCodesFiltered().get(i), "ALL")).getResultList());
        }

        getAccountsWithSubsidiaryData().setRunComplete(true);

        FacesContext.getCurrentInstance().addMessage(null, new FacesMessage(FacesMessage.SEVERITY_INFO, "Run complete", ""));
    }

    public void accountCodeMe() {
        getAccountsWithSubsidiaryData().setSubtotalPrint(BigDecimal.ZERO);

        setClassTdSummaryPrint(getCustomEntityManagerFactory().getFinancialDbEntityManagerFactory().createEntityManager().createNativeQuery(queryGenerator(getAccountsWithSubsidiaryData().getSelectedAccountCode(), "ALL")).getResultList());

        getAccountsWithSubsidiaryData().setNoOfRowsPrint(getClassTdSummaryPrint().size());

        for (int i = 0; i < getClassTdSummaryPrint().size(); i++) {
            getAccountsWithSubsidiaryData().setSubtotalPrint(getAccountsWithSubsidiaryData().getSubtotalPrint().add(((BigDecimal) getClassTdSummaryPrint().get(i)[5])));
        }

        getPrintController().method0(getClassTdSummaryPrint().size());
    }

    public void batchMe() {
        try {
            setClassTdSummaryPrint(getCustomEntityManagerFactory().getFinancialDbEntityManagerFactory().createEntityManager().createNativeQuery(queryGenerator(getAccountsWithSubsidiaryData().getSelectedAccountCode(), getPrintController().getSelectedBatch())).getResultList());
        } catch (Exception e) {

        }
    }

    public void computeTotal() {
        String query;

        for (int i = 0; i < getAccountsWithSubsidiaryData().getAccountCodesFiltered().size(); i++) {
            query = "FROM (SELECT DISTINCT ON (apd.acctno) sl.td_sl_no AS slNo, "
                    + "apd.acctno AS acctno, "
                    + "apd.account_code AS accountCode, "
                    + "cls.acct_name AS acctName, "
                    + "sl.acct_balance AS balance, "
                    + "apd.acct_create_date AS acctCreateDate, "
                    + "sl.post_date AS postDate, "
                    //                    new code below
                    + "apd.acct_status AS acctStatus, "
                    //                    new code above
                    + "cls.ctd_no AS ctdNo "
                    + "FROM coop_fin_acct_profile_dtl apd "
                    + "JOIN coop_fin_class_td cls ON apd.acctno::text = cls.acctno::text "
                    + "JOIN coop_fin_td_sl sl ON cls.acctno::text = sl.acctno::text ";

            if (getAccountsWithSubsidiaryData().getAcctCreateDateFrom() != null && getAccountsWithSubsidiaryData().getAcctCreateDateTo() != null) {
                query += "WHERE apd.acct_create_date BETWEEN '" + getCustomDate().formatDate(getAccountsWithSubsidiaryData().getAcctCreateDateFrom(), "yyyy-MM-dd") + "' AND '" + getCustomDate().formatDate(getAccountsWithSubsidiaryData().getAcctCreateDateTo(), "yyyy-MM-dd") + "' ";
            } else if (getAccountsWithSubsidiaryData().getAcctCreateDateFrom() != null && getAccountsWithSubsidiaryData().getAcctCreateDateTo() == null) {
                query += "WHERE apd.acct_create_date BETWEEN '" + getCustomDate().formatDate(getAccountsWithSubsidiaryData().getAcctCreateDateFrom(), "yyyy-MM-dd") + "' AND '" + getCustomDate().formatDate(getCustomDate().getCurrentDate(), "yyyy-MM-dd") + "' ";
            } else if (getAccountsWithSubsidiaryData().getAcctCreateDateFrom() == null && getAccountsWithSubsidiaryData().getAcctCreateDateTo() == null) {
                query += "WHERE apd.acct_create_date BETWEEN '1900-01-01' AND '" + getCustomDate().formatDate(getCustomDate().getCurrentDate(), "yyyy-MM-dd") + "' ";
            }

            if (getAccountsWithSubsidiaryData().getReportDate() == null) {
                query += "AND sl.post_date <= '" + getCustomDate().formatDate(getCustomDate().getCurrentDate(), "yyyy-MM-dd") + "' ";
            } else {
                query += "AND sl.post_date <= '" + getCustomDate().formatDate(getAccountsWithSubsidiaryData().getReportDate(), "yyyy-MM-dd") + "' ";
            }

            query += "AND apd.account_code = '" + getAccountsWithSubsidiaryData().getAccountCodesFiltered().get(i) + "' ";

            if (getAccountsWithSubsidiaryData().getAmountFilter() != null && getAccountsWithSubsidiaryData().getAmountFilter().compareTo(BigDecimal.ZERO) == 1) {
                query += "AND sl.acct_balance <= " + getAccountsWithSubsidiaryData().getAmountFilter();
            }

//            new code below
//            query += "AND (apd.acct_status <> 'C') "
//                    + "OR (apd.acct_status = 'C' "
//                    + "AND sl.acct_balance <> 0) ";
//            new code above   
            query += " ORDER BY apd.acctno, sl.post_date DESC, sl.td_sl_no DESC) cls "
                    + "WHERE (cls.acctStatus = 'A') "
                    + "OR (cls.acctStatus = 'C' "
                    + "AND cls.balance <> 0) "
                    + "OR (cls.acctStatus = 'O' "
                    + "AND cls.balance <> 0) ";
//                    + "ORDER BY cls.acctno";

            getAccountsWithSubsidiaryData().getNoOfRows().add(i, (Long) getCustomEntityManagerFactory().getFinancialDbEntityManagerFactory().createEntityManager().createNativeQuery("SELECT COUNT(cls) " + query).getResultList().get(0));

            getAccountsWithSubsidiaryData().getSubtotal().add(i, (BigDecimal) getCustomEntityManagerFactory().getFinancialDbEntityManagerFactory().createEntityManager().createNativeQuery("SELECT COALESCE(SUM(cls.balance), 0) " + query).getResultList().get(0));

            getAccountsWithSubsidiaryData().setGrandTotal(getAccountsWithSubsidiaryData().getGrandTotal().add(getAccountsWithSubsidiaryData().getSubtotal().get(i)));
        }
    }

    public void export0() {

        Integer columnNo;
        HSSFWorkbook workbook;
        HSSFSheet sheet;
        HSSFRow headerRow, dataRow, totalRow = null;
        HSSFCell cell;
        HSSFCellStyle cellStyle, boldStyle;
        HSSFFont font;

        ThemeDisplay themeDisplay = LiferayFacesContext.getInstance().getThemeDisplay();

        getExportData().createFolder(null, themeDisplay, "Time Deposit Report", "DESCRIPTION");

        if (getExportData().getFilename() == null || getExportData().getFilename().length() == 0) {
            getExportData().setFilename("Default(" + new Date() + ")");
        }

        try {
            getExportData().setFilename(getExportData().getFilename().replace(":", ""));
            getExportData().setFilename(getExportData().getFilename().concat(".xls"));
            workbook = new HSSFWorkbook();

            cellStyle = workbook.createCellStyle();
            cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));

            font = workbook.createFont();
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);

            boldStyle = workbook.createCellStyle();
            boldStyle.setFont(font);

            for (int i = 0; i < getAccountsWithSubsidiaryData().getAccountCodesFiltered().size(); i++) {

                try {
                    sheet = workbook.createSheet(getDataConvert().accountCodeConvert(getAccountsWithSubsidiaryData().getAccountCodesFiltered().get(i)));
                } catch (Exception e) {
                    sheet = workbook.createSheet(getAccountsWithSubsidiaryData().getAccountCodesFiltered().get(i));
                }

                headerRow = sheet.createRow((short) 0);
                columnNo = 0;
                cell = headerRow.createCell(columnNo++);
                cell.setCellValue(getDataConvert().accountCodeConvert(getAccountsWithSubsidiaryData().getAccountCodesFiltered().get(i)));
                cell.setCellStyle(boldStyle);

                try {
                    headerRow = sheet.createRow(headerRow.getRowNum() + 1);
                    columnNo = 0;
                    cell = headerRow.createCell(columnNo++);
                    cell.setCellValue("As of " + getCustomDate().formatDate(getAccountsWithSubsidiaryData().getReportDate(), "MM-dd-YYYY"));
                    cell.setCellStyle(boldStyle);
                } catch (Exception e) {
                    getAccountsWithSubsidiaryData().setReportDate(getCustomDate().getCurrentDate());

                    cell.setCellValue("As of " + getCustomDate().formatDate(getAccountsWithSubsidiaryData().getReportDate(), "MM-dd-YYYY"));
                    cell.setCellStyle(boldStyle);
                }

                if (getAccountsWithSubsidiaryData().getAcctCreateDateFrom() != null
                        || getAccountsWithSubsidiaryData().getAcctCreateDateTo() != null) {
                    headerRow = sheet.createRow(headerRow.getRowNum() + 1);
                    columnNo = 0;
                    cell = headerRow.createCell(columnNo++);
                    cell.setCellValue(getAccountsWithSubsidiaryData().getAcctCreateDateTo() != null ? "Account Created Date: " + getCustomDate().formatDate(getAccountsWithSubsidiaryData().getAcctCreateDateFrom(), "MM-dd-YYYY").concat(" - ").concat(getCustomDate().formatDate(getAccountsWithSubsidiaryData().getAcctCreateDateTo(), "MM-dd-YYYY")) : "Account Created Date: " + getCustomDate().formatDate(getAccountsWithSubsidiaryData().getAcctCreateDateFrom(), "MM-dd-YYYY"));
                    cell.setCellStyle(boldStyle);
                }

                if ((getAccountsWithSubsidiaryData().getAmountFilter() != null && getAccountsWithSubsidiaryData().getAmountFilter().compareTo(BigDecimal.ZERO) == 1)) {
                    headerRow = sheet.createRow(headerRow.getRowNum() + 1);
                    columnNo = 0;
                    cell = headerRow.createCell(columnNo++);
                    cell.setCellValue("Amount Range: " + getDataConvert().numericConvert(getAccountsWithSubsidiaryData().getAmountFilter().doubleValue()));
                    cell.setCellStyle(boldStyle);
                }

                headerRow = sheet.createRow(headerRow.getRowNum() + 1);
                headerRow = sheet.createRow(headerRow.getRowNum() + 1);
                columnNo = 0;

                cell = headerRow.createCell(columnNo++);
                cell.setCellValue("Account No.");
                cell.setCellStyle(boldStyle);

                cell = headerRow.createCell(columnNo++);
                cell.setCellValue("Name");
                cell.setCellStyle(boldStyle);

                cell = headerRow.createCell(columnNo++);
                cell.setCellValue("Certificate No.");
                cell.setCellStyle(boldStyle);

                cell = headerRow.createCell(columnNo++);
                cell.setCellValue("Account Status");
                cell.setCellStyle(boldStyle);

                cell = headerRow.createCell(columnNo++);
                cell.setCellValue("Balance");
                cell.setCellStyle(boldStyle);

                //add space?
                cell = headerRow.createCell(columnNo++);
                cell.setCellValue(" ");

                cell = headerRow.createCell(columnNo++);
                cell.setCellValue("SC Account no.");
                cell.setCellStyle(boldStyle);

                for (int ii = 0; ii < getClassTdSummary().get(i).size(); ii++) {

                    columnNo = 0;
                    dataRow = sheet.createRow(headerRow.getRowNum() + ii + 1);
                    //acctno
                    try {
                        String cellD = ((String) getClassTdSummary().get(i).get(ii)[2]);
                        System.out.println("CHEKING acctno IF NULL: " + cellD);
                        if (cellD != null) {
                            dataRow.createCell(columnNo++).setCellValue(cellD);
                        }
                    } catch (Exception e) {
                        System.out.println("Error on cellD " + e.getMessage());
                        dataRow.createCell(columnNo++).setCellValue("-");
                    }

                    //acctname
                    try {
                        String cellAcctName = ((String) getClassTdSummary().get(i).get(ii)[4].toString());
                        System.out.println("ACCT NAME ITO: " + cellAcctName);
                        dataRow.createCell(columnNo++).setCellValue(cellAcctName);
                    } catch (Exception e) {
                        System.out.println("Error on acctName " + e.getMessage());
                        dataRow.createCell(columnNo++).setCellValue("-");
                    }

                    //ctd no.
                    try {
                        Integer ctdNo = ((Integer) getClassTdSummary().get(i).get(ii)[9]);
                        System.out.println("CHEKING ctdNo IF NULL: " + ctdNo);
                        if (ctdNo != null) {
                            dataRow.createCell(columnNo++).setCellValue(ctdNo);
                        } else {

                            System.out.println("Error: Certificate No. is not an Integer");
                            dataRow.createCell(columnNo++).setCellValue("-");
                        }
                    } catch (Exception e) {
                        System.out.println("error on ctdNo. " + e.getMessage());
                        dataRow.createCell(columnNo++).setCellValue("-");
                    }

                    //acctstatus
                    try {
                        dataRow.createCell(columnNo++).setCellValue(getDataConvert().acctStatusConvert((Character) getClassTdSummary().get(i).get(ii)[6].toString().charAt(0)));
                        cell = dataRow.createCell(columnNo++);
                        cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                    } catch (Exception e) {
                        System.out.println("Error on acctStatus " + e.getMessage());
                    }

                    //balance
                    try {
                        BigDecimal cellBigD = ((BigDecimal) getClassTdSummary().get(i).get(ii)[5]);

                        System.out.println("CHEKING balance IF NULL: " + cellBigD);
                        if (cellBigD != null) {
                            dataRow.createCell(columnNo++).setCellValue(cellBigD.doubleValue());
                        } else {
                            dataRow.createCell(columnNo++).setCellValue("-");
                        }
                    } catch (Exception e) {
                        System.out.println("Error on Balance " + e.getMessage());
                        dataRow.createCell(columnNo++).setCellValue("-");
                    }

                    //primary holder  -not needed               
//                    String cellPrime = getClassTdSummary().get(i).get(ii)[10].toString();
//                    System.out.println("CHEKING IF NULL: " + cellPrime);
//                    cell = dataRow.createCell(columnNo++);
//                    try {
//                        if (!cellPrime.isEmpty()) {
//                            cell.setCellValue(cellPrime);
//                        }
//                    } catch (Exception e) {
//                        cell.setCellValue("-");
//                        System.out.println("Error on primary holder " + e);
//                    }

                    /*
                     revised
                     sc_acctno
                     */
                    try {
                        String cellScAcctno = getClassTdSummary().get(i).get(ii)[11].toString();
                        System.out.println("CHEKING sc_acctno IF NULL: " + cellScAcctno);
                        if (cellScAcctno != null) {
                            dataRow.createCell(columnNo++).setCellValue(cellScAcctno);
                        } else {
                            dataRow.createCell(columnNo++).setCellValue("No SC");
                        }
                    } catch (Exception e) {

                        dataRow.createCell(columnNo++).setCellValue("No Sc");
                        System.out.println("Error on scAcctno " + e.getMessage());
                    }

                    cell.setCellStyle(cellStyle);
                    totalRow = sheet.createRow((short) dataRow.getRowNum() + 2);
                }

                if (getClassTdSummary().get(i).size() > 0) {

                    cell = totalRow.createCell(1);
                    cell.setCellValue("TOTAL");
                    cell.setCellStyle(boldStyle);

                    cell = totalRow.createCell(5);
                    cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);

                    Double cellDouble = getAccountsWithSubsidiaryData().getSubtotal().get(i).doubleValue();
                    if (cellDouble != null) {
                        cell.setCellValue(cellDouble);
                    } else {
                        cell.setCellValue("-");
                    }
                    cell.setCellStyle(cellStyle);
                }
            }
            //export
            try {
                FileOutputStream fileOutputStream = new FileOutputStream(getExportData().getFilename());
                workbook.write(fileOutputStream);
                fileOutputStream.close();
                getExportData().fileUploadByDL(getExportData().getFilename(), "Time Deposit Report", themeDisplay, null);

                File file = new File(getExportData().getFilename());
                if (file.exists()) {
                    file.delete();
                }
            } catch (Exception e) {
                e.getMessage();
                FacesMessage message = new FacesMessage(FacesMessage.SEVERITY_ERROR, "Error", "An error occurred while generating excel file.");
                FacesContext.getCurrentInstance().addMessage(null, message);
            }

        } catch (Exception e) {
            System.out.print("classTdData().export0() " + e);
            FacesMessage message = new FacesMessage(FacesMessage.SEVERITY_ERROR, "Error", "An error occurred while generating excel file.");
            FacesContext.getCurrentInstance().addMessage(null, message);
        }
    }

}
