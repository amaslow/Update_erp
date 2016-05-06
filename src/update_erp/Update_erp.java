package update_erp;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

class Update_erp {

    static String[] brandList = {"", "RANEX", "XQLITE", "IGLOW", "GAMMA", "MIRO",
        "ELRO", "LIEF", "SMARTWARES", "ALPHA", "DUOLEC", "HOMEWIZARD", "HOMEBASE",
        "EATEL", "NEDIS", "KRUIDVAT", "TREKPLEISTER", "WATSHOME", "KWANTUM", "KARWEI",
        "HEMA", "CASALUX", "OK", "DIFFERENZ", "INTERTOYS", "KONZUM", "NEXT", "PALATRADE", "CASAYA", "ANSLUT"};

    public static void main(String[] args) throws IOException, ParseException {

        String excelname = "G:\\Share Company Wide\\Company Transfer\\ERP classificatie\\ERP bulbs.xlsx";
        Connection con = null;
        Statement st = null;
        ResultSet rs = null;
        String sap_db, item_db, ean_db, enclas_db, brand_db;
        String sap_xlsx, item_xlsx, ean_xlsx, enclas_xlsx, brand_xlsx;
        double wattage_xlsx, wattage_db;
        int brandN_new = 0;
        
        FileWriter fw = new FileWriter("H:/Logs/UpdateErp.log", true);
        BufferedWriter bw = new BufferedWriter(fw);
        DateFormat dateFormater = new SimpleDateFormat("dd.MM.yyyy");
        String modDate = dateFormater.format(new Date());
        bw.newLine();
        bw.write(modDate);
        
        FileInputStream fis = null;
        fis = new FileInputStream(new File(excelname));
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);

        try {
            con = Utils.getConnection();
            st = con.createStatement();

            String SQL_UPDATE = "UPDATE elro.items SET enclas='NA' where (enclas is null or enclas='') and KIND_BULB is not null and KIND_BULB<>'' ORDER BY sap;";
            st.executeUpdate(SQL_UPDATE);

            String SQL_STATUS = "SELECT sap,item,ean,wattage,enclas,brand FROM elro.items where KIND_BULB<>'Luminaire' and KIND_BULB is not null and KIND_BULB<>'' ORDER BY sap;";
            rs = st.executeQuery(SQL_STATUS);

            while (rs.next()) {
                sap_db = rs.getString(1);
                item_db = rs.getString(2);
                ean_db = rs.getString(3);
                wattage_db = rs.getDouble(4);
                enclas_db = rs.getString(5);
                brand_db = rs.getString(6).replace("PL ", "");

                int rownr = findRow(sheet, sap_db);
                if (rownr != 0) {

                    XSSFRow row = sheet.getRow(rownr);
                    XSSFCell item_xlsx1 = row.getCell(0);
                    item_xlsx = item_xlsx1.getStringCellValue();
                    if (!item_db.equals(item_xlsx)) {
                        System.out.println("Change item for : " + sap_db + " - new: " + item_db + " /old: " + item_xlsx);
                        bw.newLine();
                        bw.write("\t" + "Change item for : " + sap_db + " - new: " + item_db + " /old: " + item_xlsx);
                        item_xlsx1.setCellValue(item_db);
                    }
                    XSSFCell ean_xlsx1 = row.getCell(2);
                    ean_xlsx = ean_xlsx1.getStringCellValue();
                    if (!ean_db.equals(ean_xlsx)) {
                        System.out.println("Change ean for : " + sap_db + " - new: " + ean_db + " /old: " + ean_xlsx);
                        bw.newLine();
                        bw.write("\t" + "Change ean for : " + sap_db + " - new: " + ean_db + " /old: " + ean_xlsx);
                        ean_xlsx1.setCellValue(ean_db);
                    }
                    XSSFCell wattage_xlsx1 = row.getCell(3);
                    wattage_xlsx = wattage_xlsx1.getNumericCellValue();
                    if (Math.round(wattage_db) != wattage_xlsx) {
                        System.out.println("Change wattage for : " + sap_db + " - new: " + wattage_db + " /old: " + wattage_xlsx);
                        bw.newLine();
                        bw.write("\t" + "Change wattage for : " + sap_db + " - new: " + wattage_db + " /old: " + wattage_xlsx);
                        wattage_xlsx1.setCellValue(Math.round(wattage_db));
                    }
                    XSSFCell enclas_xlsx1 = row.getCell(4);
                    enclas_xlsx = enclas_xlsx1.getStringCellValue();
                    if (!enclas_db.equals(enclas_xlsx)) {
                        System.out.println("Change enclas for : " + sap_db + " - new: " + enclas_db + " /old: " + enclas_xlsx);
                        bw.newLine();
                        bw.write("\t" + "Change enclas for : " + sap_db + " - new: " + enclas_db + " /old: " + enclas_xlsx);
                        enclas_xlsx1.setCellValue(enclas_db);
                    }
                    XSSFCell brand_xlsx1 = row.getCell(5);
                    brand_xlsx = brand_xlsx1.getStringCellValue();
                    int brandN = Integer.parseInt(brand_xlsx.substring(5));

                    for (int i = 0; i < brandList.length; i++) {
                        if (brandList[i].equals(brand_db)) {
                            brandN_new = i;
                            break;
                        }
                    }
                    if (brand_db.equals("UNBRANDED") || brand_db.equals("ALDI") || brand_db.equals("BYRON")
                            || brand_db.equals("NORMA")) {
                        brand_db = "";
                        brandN_new = 0;
                    }
                    if (brand_db.equals("XQ-LITE BY COSM")) {
                        brand_db = "XQLITE";
                        brandN_new = 2;
                    }
                    if (brand_db.equals("GAMMA NL") || brand_db.equals("GAMMA BE")
                            || brand_db.equals("GAMMA / JDB") || brand_db.equals("GAMMA / OK")) {
                        brand_db = "GAMMA";
                        brandN_new = 4;
                    }
                    String brand_xlsx11 = brandList[brandN].toString();
                    if (!brand_db.equals(brand_xlsx11) || brandN != brandN_new) {
                        System.out.println("Change brand for : " + sap_db + " - new: " + brand_db + "(" + brandN_new + ")" + " /old: " + brand_xlsx11 + "(" + brandN + ")");
                        bw.newLine();
                        bw.write("\t" + "Change brand for : " + sap_db + " - new: " + brand_db + "(" + brandN_new + ")" + " /old: " + brand_xlsx11 + "(" + brandN + ")");
                        brand_xlsx1.setCellValue("Logo_" + brandN_new);
                    }

                } else {
                    for (int i = 0; i < brandList.length; i++) {
                        if (brandList[i].equals(brand_db)) {
                            brandN_new = i;
                            break;
                        }
                    }
                    System.out.println("New item: " + item_db + " | " + sap_db + " | " + ean_db + " | " + wattage_db + " | " + enclas_db + " | " + brand_db + "(" + brandN_new + ")");
                    bw.newLine();
                    bw.write("\t" + "New item: " + item_db + " | " + sap_db + " | " + ean_db + " | " + wattage_db + " | " + enclas_db + " | " + brand_db + "(" + brandN_new + ")");
                    int lastRow = sheet.getLastRowNum() + 1;
                    XSSFRow rowNew = sheet.createRow(lastRow);
                    XSSFCell itemCell = rowNew.createCell(0);
                    itemCell.setCellValue(item_db);
                    XSSFCell sapCell = rowNew.createCell(1);
                    sapCell.setCellValue(sap_db);
                    XSSFCell eanCell = rowNew.createCell(2);
                    eanCell.setCellValue(ean_db);
                    XSSFCell wattageCell = rowNew.createCell(3);
                    wattageCell.setCellValue(Math.round(wattage_db));
                    XSSFCell enclasCell = rowNew.createCell(4);
                    enclasCell.setCellValue(enclas_db);
                    XSSFCell brandCell = rowNew.createCell(5);
                    brandCell.setCellValue("Logo_" + brandN_new);
                }
            }
        } catch (SQLException ex) {
            Logger.getLogger(Update_erp.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            Utils.closeDB(rs, st, con);
        }
        fis.close();
        FileOutputStream fos = new FileOutputStream(new File(excelname));
        wb.write(fos);
        fos.close();
        bw.newLine();
        bw.write("----------------------------------------------------");
        bw.flush();
        bw.close();
    }

    private static int findRow(XSSFSheet sheet, String item) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    if (cell.getRichStringCellValue().getString().trim().equals(item)) {
                        return row.getRowNum();
                    }
                }
            }
        }
        return 0;
    }
}
