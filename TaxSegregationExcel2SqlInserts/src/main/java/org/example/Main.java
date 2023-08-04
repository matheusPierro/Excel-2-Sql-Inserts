package org.example;

import org.apache.poi.ss.usermodel.*;
import java.io.*;

public class Main {
    public static void main(String[] args) {
        File f = new File("C:\\Users\\matheus.r.pierro\\Downloads\\SmartHistory_T511896_CFG511894_SegregacaoReceita_v1.7.xlsx");
        File file = new File("C:\\Users\\matheus.r.pierro\\Downloads\\inserts\\input.sql");
        BufferedWriter bw = null;

        try {
            bw = new BufferedWriter(new FileWriter(file));

            Workbook workbook = WorkbookFactory.create(f);
            Sheet sheet = workbook.getSheetAt(1);

            String prev_area_codes = "";
            String prev_product_name = "";
            String prev_ext_id = "";
            double prev_seg_value = 0.0;
            double prev_prod_value = 0.0;
            String prev_reason = "";
            String prev_step = "";
            String prev_efficiency = "";

            int c = 0;

            for (Row row : sheet) {
                c++;
                if (c > 3) {
                    Cell area_codeCell = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    Cell product_nameCell = row.getCell(2, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    Cell product_extidCell = row.getCell(3, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    Cell product_seg_nameCell = row.getCell(6, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    Cell companyCell = row.getCell(7, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    Cell product_seg_valueCell = row.getCell(8, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    Cell product_valueCell = row.getCell(5, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    Cell stepCell = row.getCell(4, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    Cell reasonCell = row.getCell(11, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    Cell efficiencyCell = row.getCell(1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

                    System.out.println(product_seg_nameCell);

                    String area_code = (area_codeCell == null) ? prev_area_codes : area_codeCell.toString();
                    String product_name = (product_nameCell == null) ? prev_product_name : product_nameCell.toString();
                    String product_extid = (product_extidCell == null) ? prev_ext_id : product_extidCell.toString();
                    double product_seg_value = (product_seg_valueCell == null || product_seg_valueCell.getCellType() != CellType.NUMERIC) ? prev_seg_value : product_seg_valueCell.getNumericCellValue();
                    double product_value = (product_valueCell == null || product_valueCell.getCellType() != CellType.NUMERIC) ? prev_prod_value : product_valueCell.getNumericCellValue();
                    String efficiency = (efficiencyCell == null) ? prev_efficiency : efficiencyCell.toString();
                    efficiency = (efficiency.toUpperCase().equals("SIM")) ? "Y" : "N";
                    String reason = (reasonCell == null || reasonCell.getCellType() != CellType.NUMERIC) ? "null" : String.format("%d", (int) Double.parseDouble(reasonCell.toString().trim()));

                    String step = (stepCell == null) ? prev_step : stepCell.toString();

                    if (!product_seg_nameCell.toString().equals("Produto")) {
                        String[] area_codes = area_code.replaceAll("\\s", ",").replaceAll(",,", ",").split(",");

                        for (String code : area_codes) {
                            String query = String.format(
                                    "UPSERT INTO SMARTHISTORY_LAB5.FIN_CFG_TAX_SEGREGATION(AREA_CODE,PRODUCT_NAME,PRODUCT_EXTERNAL_ID,SEG_PRODUCT_NAME,COMPANY,SEG_VALUE,SEG_TYPE,PRODUCT_VALUE,REASON_RM,PRIORITY,STEP, EFFICIENCY) VALUES(%d,'%s','%s','%s','%s',%.2f,'PERC',%.2f,%s,%d,'%s','%s');\n",
                                    (int) Double.parseDouble(code.trim()),
                                    product_name,
                                    product_extid,
                                    product_seg_nameCell.toString(),
                                    companyCell.toString(),
                                    product_seg_value,
                                    product_value,
                                    reason,
                                    1,
                                    step,
                                    efficiency
                            );

                            System.out.print(query);
                            bw.write(query);
                        }
                    }

                    prev_area_codes = (area_codeCell == null) ? prev_area_codes : area_codeCell.toString();
                    prev_product_name = (product_nameCell == null) ? prev_product_name : product_nameCell.toString();
                    prev_ext_id = (product_extidCell == null) ? prev_ext_id : product_extidCell.toString();
                    prev_seg_value = (product_seg_valueCell == null || product_seg_valueCell.getCellType() != CellType.NUMERIC) ? prev_seg_value : product_seg_valueCell.getNumericCellValue();
                    prev_prod_value = (product_valueCell == null || product_valueCell.getCellType() != CellType.NUMERIC) ? prev_prod_value : product_valueCell.getNumericCellValue();
                    prev_efficiency = (efficiencyCell == null) ? prev_efficiency : efficiencyCell.toString();
                    prev_step = (stepCell == null) ? prev_step : stepCell.toString();
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (bw != null) {
                try {
                    bw.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
