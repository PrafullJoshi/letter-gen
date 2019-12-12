package in.co.parna.lettergen.service;

import com.lowagie.text.*;
import com.lowagie.text.Font;
import com.lowagie.text.Image;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import in.co.parna.lettergen.dto.LetterGeneratorData;
import in.co.parna.lettergen.dto.LetterGeneratorRequestDto;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Iterator;

@Service
public class LetterGeneratorService {

    @Value("${letter.generator.inputFilePathname}")
    String inputFilePathname;

    @Value("${letter.generator.logoPath}")
    String logoPath;

    @Value("${letter.generator.document.count.threshold}")
    int threshold;

    private static String OUTPUT_FILE = "output/#_Madhukosh_Maintainance.pdf";

    // COLORs
    private static final Color logoColor = new Color(4, 120, 59);
    private static final Color hiveColor = new Color(252, 196, 38);
    private static final Color primaryColor = new Color(253, 180, 7);
    private static final Color secondaryColor = new Color(91, 255, 255);
    private static final Color tertiaryColor = new Color(139, 255, 83);

    // FONTS
    private static Font regularFont = new Font(Font.HELVETICA, 8, Font.NORMAL);
    private static Font regularUnderlinedFont = new Font(Font.HELVETICA, 8, Font.UNDERLINE);
    private static Font regularFontBold = new Font(Font.HELVETICA, 8, Font.BOLD);
    private static Font regularFontBoldWhite = new Font(Font.HELVETICA, 8, Font.BOLD, Color.white);
    private static Font regularFontBoldItalic = new Font(Font.HELVETICA, 8, Font.BOLDITALIC);
    private static Font regularFontBoldLogoColor = new Font(Font.HELVETICA, 8, Font.BOLD, logoColor);


    private static Font smallFont = new Font(Font.HELVETICA, 8, Font.NORMAL);
    private static Font xxsmallFont = new Font(Font.HELVETICA, 2, Font.NORMAL);
    private static Font emailFont = new Font(Font.HELVETICA, 8, Font.UNDERLINE, Color.BLUE);

    public void generateLetters(LetterGeneratorRequestDto letterGeneratorRequestDto) throws IOException, DocumentException {

        FileInputStream excelFile = new FileInputStream(new File(inputFilePathname));

        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = datatypeSheet.iterator();

        int counter = 0;

        while (iterator.hasNext()) {

            org.apache.poi.ss.usermodel.Row currentRow = iterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();

            if(currentRow.getRowNum() == 0) {
                continue;
            }
            if(counter == threshold) {
                break;
            }

            LetterGeneratorData data = new LetterGeneratorData();

            while (cellIterator.hasNext()) {

                Cell currentCell = cellIterator.next();
                switch(currentCell.getColumnIndex()) {
                    case 0 : // Flat No A
                        if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                            data.setFlatNo(currentCell.getStringCellValue());
                        } else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC || currentCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                            // currentName = currentCell.getNumericCellValue();
                        }
                        break;
                    case 1 : // Name B
                        if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                            data.setOwner(currentCell.getStringCellValue());
                        } else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC || currentCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                            // currentName = currentCell.getNumericCellValue();
                        }
                        break;
                    case 2 : // Area C
                        if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                            data.setAreaSquareFeet(currentCell.getStringCellValue());
                        } else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC || currentCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                                double numericCellValue = currentCell.getNumericCellValue();
                                String format = String.format("%.0f", numericCellValue);
                                data.setAreaSquareFeet(String.valueOf(numericCellValue));
                         }
                        break;
                    case 3 : // Advance D
                        double numericCellValue = currentCell.getNumericCellValue();
                        String format = String.format("%.0f", numericCellValue);
                        data.setAdvance(format);
                        /*if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                            data.setAdvance(currentCell.getStringCellValue());
                        }*/
                        break;
                    case 5 : // F
                        if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                            data.setLateFees18_19(currentCell.getStringCellValue());
                        } else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC || currentCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                            data.setLateFees18_19(String.valueOf(currentCell.getNumericCellValue()));
                        }

                        break;
                    case 7 : // H
                        if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                            data.setLateFees19_20(currentCell.getStringCellValue());
                        } else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC || currentCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                            data.setLateFees19_20(String.valueOf(currentCell.getNumericCellValue()));
                        }
                        break;
                    case 9 : // J
                        if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                            data.setTotalLateFees(currentCell.getStringCellValue());
                        } else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC || currentCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                            data.setTotalLateFees(String.valueOf(currentCell.getNumericCellValue()));
                        }
                        break;
                    case 10 : // K
                        if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                            data.setPreviousOutstandingDues(currentCell.getStringCellValue());
                        } else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC || currentCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                            data.setPreviousOutstandingDues(String.format("%.0f", currentCell.getNumericCellValue()));
                        }
                        break;
                    case 13 : // N
                        if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                            data.setPerAreaSqFeetExpenses(currentCell.getStringCellValue());
                        } else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC || currentCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                            data.setPerAreaSqFeetExpenses(String.format("%.0f", currentCell.getNumericCellValue()));
                        }
                        break;
                    case 16 : //Q
                        if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                            data.setPromptDiscount18_19(currentCell.getStringCellValue());
                        } else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC || currentCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                            data.setPromptDiscount18_19(String.format("%.0f", currentCell.getNumericCellValue()));
                        }
                        break;
                    case 17 : //R
                        if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
                            data.setPromptDiscount19_20(currentCell.getStringCellValue());
                        } else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC || currentCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                            data.setPromptDiscount19_20(String.format("%.0f", currentCell.getNumericCellValue()));
                        }
                        // data.setInterestDate(new SimpleDateFormat("dd-MMM-yy").format(currentCell.getDateCellValue()));
                        break;
                    case 18 : //S
                        data.setPromptDiscount(String.format("%.0f", currentCell.getNumericCellValue()));
                        // data.setPromptDiscountPercent1(String.format("%.0f", (currentCell.getNumericCellValue() * 100)));
                                     /*double numericCellValue = currentCell.getNumericCellValue();
                                     String format = String.format("%.0f", numericCellValue);
                                     data.setDevelopmentFundPercentage(format);*/
                        break;
                    case 20 : //U
                        data.setDevFundContri(String.format("%.0f", (currentCell.getNumericCellValue())));
                        break;
                    case 21 : // V
                        data.setTenantCharges(String.format("%.0f", currentCell.getNumericCellValue()));
                        break;
                    case 22 : // W
                        data.setAnnualDemand(String.format("%.0f", currentCell.getNumericCellValue()));
                        break;
                    case 23 : //X
                        data.setTotalAmountDue(String.format("%.0f", currentCell.getNumericCellValue()));
                        break;
                }

            }

            int TOP = 0;
            int BOTTOM = 0;
            int LEFT = 4;
            int RIGHT = 4;

            String flatNo = data.getFlatNo();
            if(flatNo != null && !"".equals(flatNo)) {
                Document document = new Document(PageSize.A4);
                String fileName = OUTPUT_FILE.replace("#", flatNo.trim());
                PdfWriter.getInstance(document, new FileOutputStream(new File(fileName)));
                document.open();
                addMetaData(document);
                addContent(document, data);
                document.setMargins(LEFT, RIGHT, TOP, BOTTOM);
                document.close();
                System.out.println(counter + " > File generated for - " + fileName);
            } else {
                // Since no flat no present, break the loop
                break;
            }
            counter++;
        }
    }

    private void addContent(Document document, LetterGeneratorData data) throws DocumentException, IOException {


//            System.out.println(data);
        Paragraph preface = new Paragraph();

        // String path = LetterGenerator.class.getResource(FG_LETTER_LOGO).getPath();

        Image image = Image.getInstance(logoPath);
        image.scaleToFit(PageSize.A4.getWidth() - 50, PageSize.A4.getHeight()- 800);
        image.setAlignment(Element.ALIGN_CENTER);
        preface.add(image);


        PdfPTable table = new PdfPTable(new float[]{1});
        PdfPCell c1 = new PdfPCell(new Phrase("S. No. 16, 4/2, 17 (Part), (Plot 1) + 14/4B (Part), Wadgaon (Kh), Dhayari, Pune - 411 068.", smallFont));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBorder(Rectangle.BOTTOM);
        c1.setBorderColor(hiveColor);
        c1.setBorderWidth(2);
        c1.setPaddingBottom(15);
        table.setWidthPercentage(100);
        table.addCell(c1);

        SimpleDateFormat df = new SimpleDateFormat("dd-MMM-yyyy");
        PdfPCell dateCell = new PdfPCell(new Phrase(df.format(Calendar.getInstance().getTime()), regularFontBold));
        dateCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        dateCell.setBorder(Rectangle.NO_BORDER);
        table.addCell(dateCell);
        preface.add(table);


        preface.add(new Paragraph(data.getOwner(), regularFontBoldLogoColor));
        preface.add(new Paragraph(data.getFlatNo() + ", Madhukosh Ph II,", regularFont));
        preface.add(new Paragraph("Dhayari, Pune - 411068", regularFont));

        addEmptyLine(preface, 1);

        Paragraph subject = new Paragraph("     Sub: Annual Maintenance charges for the period: April–2019 to March–2020.", regularFontBoldItalic);
        subject.setAlignment(Element.ALIGN_CENTER);
        preface.add(subject);

        // addEmptyLine(preface, 1);

        Paragraph salutation = new Paragraph("Dear Member,", regularFont);
        preface.add(salutation);

        Paragraph requestPara =  new Paragraph("      You are requested to pay the annual maintenance charges for the period from 01-Apr-2019 to 31-Mar-2020 as per the following details: ", regularFont);
//            Phrase dueDateNote = new Phrase(" The due date for the payment is " + data.getDueDate() + ".", regularUnderlinedFontForLetter);
//            Phrase postNote = new Phrase(" The details of charges payable from you are as follows:", regularFontForLetter);

        preface.add(requestPara);
//            preface.add(dueDateNote);
//            preface.add(postNote);

        // First Table
        createTotalDuesTable(preface, data);
        addEmptyLine(preface, 1);

        preface.add(new Paragraph("Important:", regularFontBold));
        preface.add(new Phrase("Interest charge of 1 % per month on the ", regularFont));
        preface.add(new Phrase("\"Total Amount Payable\"", regularFontBold));
        preface.add(new Paragraph(" in the table above will be applicable from 01-Jan-2020 every month till payment.", regularFont));

//        addEmptyLine(preface, 1);

        preface.add(new Phrase("Kindly refer to the ", regularFont));
        preface.add(new Phrase(" \"Late Fee\"", regularFontBold));
        preface.add(new Phrase(" Section C in the ", regularFont));
        preface.add(new Phrase(" \"Maintenance Calculation Details\"", regularFontBold));
        preface.add(new Phrase(" provided below.  If the section mentions “TBD”, then kindly refer the note below. \n" +
                "Late fee charges (TBD) are calculated after receiving the outstanding amount till the date of receipt of outstanding amount payment.  The calculated amount will be added to  ", regularFont));
        preface.add(new Phrase(" \"Total Amount Payable\"", regularFontBold));
        preface.add(new Paragraph(" after the receipt of outstanding amount payment.", regularFont));

//        addEmptyLine(preface, 1);

        preface.add(new Phrase("If the ", regularFont));
        preface.add(new Phrase(" \"Total Amount Payable\"", regularFontBold));
        preface.add(new Phrase(" is a negative number, then it indicates that you have made excess payment to the Apartment and it will be subtracted from the next demand. ", regularFont));
        preface.add(new Paragraph(" ", regularFont));

        addEmptyLine(preface, 1);
        preface.add(new Phrase("\"Maintenance Calculation Details:\"", regularFontBold));

        createFirstATable(preface, data);
        addEmptyLine(preface, 1);

        createPreviousDuesTable_B(preface, data);
        addEmptyLine(preface, 1);

        createTable_C(preface, data);
        addEmptyLine(preface, 1);

        createTable_D(preface, data);
        addEmptyLine(preface, 1);

        createTable_E(preface, data);
        addEmptyLine(preface, 1);

        createTable_F(preface, data);
        addEmptyLine(preface, 1);

        document.add(preface);

        Paragraph preface2 = new Paragraph();

        preface2.add(new Phrase("The Cheques / pay-orders / Demand Draft should be in favour of - ", regularFont));

        preface2.add(new Phrase("\"Madhukosh Apartment Building F & G\" and payable at Pune.", regularFontBold));
        preface2.add(new Phrase("You can also make the payment through NEFT or IMPS Online payment, the details of bank are as follows:", regularFont));

        preface2.add(new Phrase(" Important:", regularFontBold));
        preface2.add(new Phrase(" For", regularFont));
        preface2.add(new Phrase(" Online Payment", regularFontBold));
        preface2.add(new Phrase(" in", regularFont));
        preface2.add(new Phrase(" description", regularFontBold));
        preface2.add(new Phrase(" field put", regularFont));
        preface2.add(new Phrase(" your flat number", regularFontBold));
        preface2.add(new Phrase(" first, then other details.", regularFont));

        // Bank Information
        createBankTable(preface2);

        addEmptyLine(preface2, 1);

        preface2.add(new Paragraph("Cheque payment can be made in our apartment office and you can get the receipt, please note that the actual credit date (not cheque date) will be considered for the prompt discount schedule and office records.", regularFont));
        addEmptyLine(preface2, 1);

        preface2.add(new Paragraph("In case of NEFT online payment, kindly inform office manager and get the payment confirmed, office manager will provide you the payment confirmation e-mail as official record, if you need a hard copy please visit the office to collect the same.", regularFont));
        addEmptyLine(preface2, 1);

        preface2.add(new Paragraph("In case of IMPS mobile banking payment, kindly inform office manager the mobile number used for making your payment and get the payment confirmed, office manager will provide you the payment confirmation e-mail as official record, " +
                "if you need a hard copy please visit the office to collect the same.  IMPS mobile banking payment is becoming very difficult to trace as members transfer from different mobile numbers and " +
                "do not contact to confirm the payment, so request members to please follow above note and co-operate.", regularFont));
        addEmptyLine(preface2, 1);





        preface2.add(new Paragraph("If you have any queries, please feel free to contact:", regularUnderlinedFont));

        PdfPTable officeInfoTable = new PdfPTable(new float[]{2,3});
        officeInfoTable.setWidthPercentage(100);
        PdfPCell cellOffice = new PdfPCell(new Phrase("Office manager contact numbers:", regularFont));
        cellOffice.setHorizontalAlignment(Element.ALIGN_LEFT);
        cellOffice.setBorder(Rectangle.NO_BORDER);
        officeInfoTable.addCell(cellOffice);

        cellOffice = new PdfPCell(new Phrase("Landline No. 020-24616090", regularFont));
        cellOffice.setHorizontalAlignment(Element.ALIGN_LEFT);
        cellOffice.setBorder(Rectangle.NO_BORDER);
        officeInfoTable.addCell(cellOffice);


        cellOffice = new PdfPCell(new Phrase("Madhukosh F & G office E-Mail:", regularFont));
        cellOffice.setHorizontalAlignment(Element.ALIGN_LEFT);
        cellOffice.setBorder(Rectangle.NO_BORDER);
        officeInfoTable.addCell(cellOffice);

        Phrase email = new Phrase("madhukoshfg@gmail.com", emailFont);
        cellOffice = new PdfPCell(email);
        cellOffice.setHorizontalAlignment(Element.ALIGN_LEFT);
        cellOffice.setBorder(Rectangle.NO_BORDER);
        officeInfoTable.addCell(cellOffice);


        /*cellOffice = new PdfPCell(new Phrase("MADHUKOSH APARTMENTS (BUILDING F&G) GSTIN:", regularFont));
        cellOffice.setHorizontalAlignment(Element.ALIGN_LEFT);
        cellOffice.setBorder(Rectangle.NO_BORDER);
        officeInfoTable.addCell(cellOffice);

        cellOffice = new PdfPCell(new Phrase("27AAEAM4231C1Z1", regularFont));
        cellOffice.setHorizontalAlignment(Element.ALIGN_LEFT);
        cellOffice.setBorder(Rectangle.NO_BORDER);
        officeInfoTable.addCell(cellOffice);*/

        preface2.add(officeInfoTable);


        addEmptyLine(preface2, 1);

        addEmptyLine(preface2, 2);
        Paragraph warmRegPara = new Paragraph("With Warm regards,", regularFont);
        preface2.add(warmRegPara);
        Paragraph thanksPara = new Paragraph("Thanking you,", regularFont);
        preface2.add(thanksPara);
        addEmptyLine(preface2, 2);


        PdfPTable signatureHoldersInfoTable = new PdfPTable(new float[]{1,1,1});
        signatureHoldersInfoTable.setWidthPercentage(100);
        PdfPCell cellOfficeSignatureHolders = new PdfPCell(new Phrase("Shrinivas Dombe", regularFontBold));
        cellOfficeSignatureHolders.setHorizontalAlignment(Element.ALIGN_LEFT);
        cellOfficeSignatureHolders.setBorder(Rectangle.NO_BORDER);
        signatureHoldersInfoTable.addCell(cellOfficeSignatureHolders);

        cellOfficeSignatureHolders = new PdfPCell(new Phrase("Chandrakant Mandape", regularFontBold));
        cellOfficeSignatureHolders.setHorizontalAlignment(Element.ALIGN_LEFT);
        cellOfficeSignatureHolders.setBorder(Rectangle.NO_BORDER);
        signatureHoldersInfoTable.addCell(cellOfficeSignatureHolders);

        cellOfficeSignatureHolders = new PdfPCell(new Phrase("Aditya Barde", regularFontBold));
        cellOfficeSignatureHolders.setHorizontalAlignment(Element.ALIGN_LEFT);
        cellOfficeSignatureHolders.setBorder(Rectangle.NO_BORDER);
        signatureHoldersInfoTable.addCell(cellOfficeSignatureHolders);

        cellOfficeSignatureHolders = new PdfPCell(new Phrase("President", regularFontBold));
        cellOfficeSignatureHolders.setHorizontalAlignment(Element.ALIGN_LEFT);
        cellOfficeSignatureHolders.setBorder(Rectangle.NO_BORDER);
        signatureHoldersInfoTable.addCell(cellOfficeSignatureHolders);

        cellOfficeSignatureHolders = new PdfPCell(new Phrase("Treasurer", regularFontBold));
        cellOfficeSignatureHolders.setHorizontalAlignment(Element.ALIGN_LEFT);
        cellOfficeSignatureHolders.setBorder(Rectangle.NO_BORDER);
        signatureHoldersInfoTable.addCell(cellOfficeSignatureHolders);

        cellOfficeSignatureHolders = new PdfPCell(new Phrase("Secretary", regularFontBold));
        cellOfficeSignatureHolders.setHorizontalAlignment(Element.ALIGN_LEFT);
        cellOfficeSignatureHolders.setBorder(Rectangle.NO_BORDER);
        signatureHoldersInfoTable.addCell(cellOfficeSignatureHolders);

        preface2.add(signatureHoldersInfoTable);

        addEmptyLine(preface2, 1);
        preface2.add(new Paragraph("[This is a computer generated letter and doesn't require signature.]", smallFont));

        document.add(preface2);
    }

    private void createTotalDuesTable(Paragraph preface, LetterGeneratorData data) throws IOException, DocumentException {

        PdfPTable table = new PdfPTable(new float[]{1,1});
        table.setWidthPercentage(100);

        PdfPCell c1 = new PdfPCell(new Phrase("Total Amount Payable", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(primaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Due Date of Payment ", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(primaryColor);
        table.addCell(c1);

        BaseFont bf = BaseFont.createFont("/Library/Fonts/Arial Unicode.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

        Font regularFontBoldLogoColorUnicode = new Font(bf, 10, Font.BOLD, logoColor);

        PdfPCell cell = new PdfPCell(new Phrase(data.getTotalAmountDue() + " ₹", regularFontBoldLogoColorUnicode));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        cell = new PdfPCell(new Phrase("31-Dec-2019", regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        preface.add(table);
    }

    private void createPreviousDuesTable_B(Paragraph preface, LetterGeneratorData data) {

        PdfPTable table = new PdfPTable(new float[]{1,6});
        table.setWidthPercentage(100);

        PdfPCell c1 = new PdfPCell(new Phrase("#", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("(B) Previous Outstanding Dues ", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);


        PdfPCell cell = new PdfPCell(new Phrase("B", regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        cell = new PdfPCell(new Phrase(data.getPreviousOutstandingDues(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        preface.add(table);
    }

    private static void createFirstATable(Paragraph preface, LetterGeneratorData data) {
        PdfPTable table = new PdfPTable(new float[]{1,1,1,1,1,1,1});
        table.setWidthPercentage(100);

        PdfPCell c1 = new PdfPCell(new Phrase("#", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Area of Flat", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Area sq. ft Expenses", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);
        table.setHeaderRows(1);

        c1 = new PdfPCell(new Phrase("Per Flat Expenses", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);
        table.setHeaderRows(1);

        c1 = new PdfPCell(new Phrase("Development Fund Charges", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Tenant Charges", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("(A) Annual Demand", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);



        PdfPCell cell = new PdfPCell(new Phrase("A", regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        cell = new PdfPCell(new Phrase(data.getAreaSquareFeet(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        cell = new PdfPCell(new Phrase(data.getPerAreaSqFeetExpenses(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        cell = new PdfPCell(new Phrase("15,347.00", regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        cell = new PdfPCell(new Phrase(data.getDevFundContri(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        cell = new PdfPCell(new Phrase(data.getTenantCharges(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        cell = new PdfPCell(new Phrase(data.getAnnualDemand(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);


        preface.add(table);
    }

    private void createTable_C(Paragraph preface, LetterGeneratorData data) {

        PdfPTable table = new PdfPTable(new float[]{1,2,2,2});
        table.setWidthPercentage(100);

        PdfPCell c1 = new PdfPCell(new Phrase("#", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Late Fee FY18-19", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Late Fee FY19-20", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("(C) Total Late Fee", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(secondaryColor);
        table.addCell(c1);


        PdfPCell cell = new PdfPCell(new Phrase("C", regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        cell = new PdfPCell(new Phrase(data.getLateFees18_19(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);
        cell = new PdfPCell(new Phrase(data.getLateFees19_20(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);
        cell = new PdfPCell(new Phrase(data.getTotalLateFees(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        preface.add(table);
    }

    private void createTable_D(Paragraph preface, LetterGeneratorData data) {

        PdfPTable table = new PdfPTable(new float[]{1,6});
        table.setWidthPercentage(100);

        PdfPCell c1 = new PdfPCell(new Phrase("#", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(tertiaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("(D) FY18-19: Advance Payment", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(tertiaryColor);
        table.addCell(c1);


        PdfPCell cell = new PdfPCell(new Phrase("D", regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        cell = new PdfPCell(new Phrase(data.getAdvance(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        preface.add(table);
    }

    private void createTable_E(Paragraph preface, LetterGeneratorData data) {

        PdfPTable table = new PdfPTable(new float[]{1,2,2,2});
        table.setWidthPercentage(100);

        PdfPCell c1 = new PdfPCell(new Phrase("#", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(tertiaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("FY18-19: Prompt Discount", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(tertiaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("FY19-20: Prompt Discount", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(tertiaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("(E) Total Prompt Discount", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(tertiaryColor);
        table.addCell(c1);


        PdfPCell cell = new PdfPCell(new Phrase("E", regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        cell = new PdfPCell(new Phrase(data.getPromptDiscount18_19(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);
        cell = new PdfPCell(new Phrase(data.getPromptDiscount19_20(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);
        cell = new PdfPCell(new Phrase(data.getPromptDiscount(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        preface.add(table);
    }

    private void createTable_F(Paragraph preface, LetterGeneratorData data) {

        PdfPTable table = new PdfPTable(new float[]{1,6});
        table.setWidthPercentage(100);

        PdfPCell c1 = new PdfPCell(new Phrase("#", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(primaryColor);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Total Amount Payable (A + B + C – D – E)", regularFontBold));
        c1.setHorizontalAlignment(Element.ALIGN_CENTER);
        c1.setBackgroundColor(primaryColor);
        table.addCell(c1);


        PdfPCell cell = new PdfPCell(new Phrase("F", regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        cell = new PdfPCell(new Phrase(data.getTotalAmountDue(), regularFont));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cell);

        preface.add(table);
    }


    private static void createBankTable(Paragraph preface) {
        PdfPTable table = new PdfPTable(new float[]{1,3});
        table.setWidthPercentage(90);
        PdfPCell c1 = new PdfPCell(new Phrase("Title of the Account:", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Madhukosh Apartment Building F and G", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Name of the Bank:", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("State Bank of India", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Branch:", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Dhayari, Pune", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Branch Address:", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("S. No. 143, Surya House, Opp: Lokmat Bhavan, Dhayri, Tal: Haveli, Pune - 411041", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Account No.:", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("34756486270", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Type of A/c:", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Current Account", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("IFSC Code:", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("SBIN0017878", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("MICR Code:", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("411002116", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Beneficiary Address:", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        c1 = new PdfPCell(new Phrase("Survey No. 4/2, Madhukosh Phase II, Vadgaon Khurd, Pune -411 041", regularFont));
        c1.setHorizontalAlignment(Element.ALIGN_LEFT);
        c1.setBorder(Rectangle.NO_BORDER);
        table.addCell(c1);

        preface.add(table);
    }


    private static void addMetaData(Document document) {
        document.addTitle("Maintenance Letter");
        document.addSubject("Software by - Prafulla Joshi, G-1005");
        document.addKeywords("Java, PDF");
        document.addAuthor("Prafulla Joshi");
        document.addCreator("Prafulla Joshi");
    }

    private static void addEmptyLine(Paragraph paragraph, int number) {
        for (int i = 0; i < number; i++) {
            paragraph.add(new Paragraph(" ", xxsmallFont));
        }
    }
}
