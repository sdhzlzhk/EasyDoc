package com.glodon.tika;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColumn;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColumns;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

/**
 * Illustartes how to recover information about a document's sections using the
 * classes defined whtin the underlying openxml4j layer.
 *
 * @author Mark B
 * @version 1.00 21st April 2011
 */
public class XWPFSectionTest {

    private static final long TO_POINTS_DIVISOR = 20;
    private static final long TO_INCHES_DIVISOR = 72;
    private static final double TO_CM_MULTIPLIER = 2.54;

    /**
     * Create an instance of the XWPFSectionTest class using the following
     * parameter(s).
     *
     * @param filename An instance of the String class that encapsulates the
     *                 path to and name of a valie Word 2007 file (.docx).
     * @throws java.io.IOException Thrown if a problem occurs in the underlying
     *                             file system.
     */
    public XWPFSectionTest(String filename) throws IOException {
        File file = null;
        FileInputStream fis = null;
        BufferedInputStream bis = null;
        XWPFDocument document = null;
        XWPFParagraph paragraph = null;
        XWPFRun run = null;
        List<XWPFParagraph> paraList = null;
        Iterator<XWPFParagraph> paraListIter = null;
        CTPPr ctPPr = null;
        CTSectPr sectPr = null;
        DecimalFormat formatter = null;

        try {
            // The DataFormat object is used simply to format the column width
            // figures for display.
            formatter = new DecimalFormat("#0.00");

            // Open the Word document.
            file = new File(filename);
            fis = new FileInputStream(file);
            bis = new BufferedInputStream(fis);
            document = new XWPFDocument(bis);

            // Get a List of the pargraphs the document contains and, from that,
            // an Iterator to step through the document one paragraph as a time.
            paraList = document.getParagraphs();
            paraListIter = paraList.iterator();

            while (paraListIter.hasNext()) {

                // Print the pargraph text to illustrate how the section
                // information is bound to a specific paragraph object.
                paragraph = paraListIter.next();

                System.out.println("Pargraph text: " +
                        paragraph.getParagraphText());

                // The section information will only be bound to the final
                // paragraph in the section. If the section information
                // is missing then the call to getPPr() will retunr a null
                // value. I suspect that it is not just the section information
                // that determines whether or not the CTPPr object will be
                // created for the paragraph but only testing on more complex
                // documents will prove or disprove this conclusion.
                ctPPr = paragraph.getCTP().getPPr();
                if (ctPPr != null) {
                    // Get the CTSectPr object that contains the information
                    // about the document section and strip (some of) the
                    // information from it.
                    sectPr = ctPPr.getSectPr();
                    if(null != sectPr){
                        this.discoverSectionInfo(sectPr, formatter);
                    }
                }
            }

            // Get the CTSectPr from the document here. This will contain the
            // information for the last or only section within the document.
            sectPr = document.getDocument().getBody().getSectPr();
            this.discoverSectionInfo(sectPr, formatter);
        } finally {
            if (bis != null) {
                bis.close();
                bis = null;
            }
        }
    }

    /**
     * Interrogates the various openxml4j objects in order to discover some of
     * the information about a specific section. Currently, all it does is to
     * disocver how many columns there are in a section along with the width
     * of each column and the size of the inter-column gap (if any). There is
     * considerably more information avaliable.
     *
     * @param sectPr    An instance of a class that implements the CTSectPr
     *                  interface and which encapsulates information about a specific
     *                  section within a Word document.
     * @param formatter An instance of the DecimalFormat class that is simply
     *                  used to prepare numeric values for diaply to the user.
     */
    private void discoverSectionInfo(CTSectPr sectPr, DecimalFormat formatter) {
        List<CTColumn> columnList = null;
        CTColumns columns = null;
        CTColumn column = null;
        BigInteger bigInteger = null;
        long widthPage = 0L;
        long widthRightMargin = 0L;
        long widthLeftMargin = 0L;
        long widthColumn = 0L;
        long widthColumnSpacing = 0L;
        long totalColumnSpacing = 0L;

        System.out.println("\n****************** Section Information. ******************");

        // Recover the width of the page along with the widths of the
        // right and left hand margins.
        widthPage = sectPr.getPgSz().getW().longValue();
        widthRightMargin = sectPr.getPgMar().getRight().longValue();
        widthLeftMargin = sectPr.getPgMar().getLeft().longValue();

        // ...and print them out.
        System.out.println("Width of page: " +
                this.convertSize(widthPage, formatter));
        System.out.println("Width right margin: " +
                this.convertSize(widthRightMargin, formatter));
        System.out.println("Width left margin: " +
                this.convertSize(widthLeftMargin, formatter));

        // If the text in the section is organised into a single
        // column - whether or not the user has explicitly set the
        // number of columns to one when they created the section -
        // the call to getCols() will return a null value. In this
        // case, the width of the column will be calculated by
        // subtracting the widths of the right and left margins from
        // the width of the page.
        columns = sectPr.getCols();
        if (columns.getNum() == null) {
            System.out.println("The text in this section is " +
                    "organised into a single column.");
            widthColumn = widthPage - (widthRightMargin + widthLeftMargin);
            System.out.println("The width of the column is: " +
                    this.convertSize(widthColumn, formatter));
        } else {

            // The section has been organised into more than one column so
            // display how many there are.
            System.out.println("The text in this section is organised into " +
                    columns.getNum().longValue() +
                    " columns.");

            // Get a List of CTColumn objects from the CTColumns object.
            /**源码*/
            //columnList = columns.getColList();
            columnList = Arrays.asList(columns.getColArray());

            // If the length of this list os zero, then all of the columns will
            // be the same size and separated by an inter-column gap whichis
            // liewise the same. In this case, it is safe to caculate the widths
            // of the columns 'manually' so to speak.
            if (columnList.size() == 0) {
                // Firstly, get the width of the inter-column space
                widthColumnSpacing = columns.getSpace().longValue();

                // If there are more than two columns, the inter-column
                // space must be totalled
                totalColumnSpacing = widthColumnSpacing *
                        (columns.getNum().longValue() - 1);

                // Now determine the width of an individual column
                // by suntracting the widths of the right and left
                // columns along with the total width of the inter-column
                // space(s) and then dividing the result by the number
                // of columns in the section.
                widthColumn = widthPage -
                        (widthRightMargin + widthLeftMargin + totalColumnSpacing);
                widthColumn = widthColumn / columns.getNum().longValue();

                // ...and then print the columns width and gap.
                System.out.println("The columns are each " +
                        this.convertSize(widthColumn, formatter) +
                        " wide.");
                System.out.println("The columns are spaced " +
                        this.convertSize(widthColumnSpacing, formatter) +
                        " apart.");
            } else {

                // If the columns list actually has CTColumn objects in it, call
                // iterateColumns() to print out the details for each.
                this.ierateColumns(columnList, formatter);
            }
        }
        System.out.println("****************** End Of Section Information. ******************\n");
    }

    /**
     * Print out the width of each column and the size of the inter-column gap.
     *
     * @param columnList A list of the columns 'contained' within a specific
     *                   section.
     * @param formatter  An inatnce of the DecimalFormat class that is used to
     *                   prepare numeric values for display to the user.
     */
    private void ierateColumns(List<CTColumn> columnList, DecimalFormat formatter) {
        CTColumn column = null;
        Iterator<CTColumn> columnListIter = null;
        BigInteger bigInteger = null;
        columnListIter = columnList.iterator();

        // Simply iterate through the columns and print out the width and
        // inter-column gap for each.
        while (columnListIter.hasNext()) {
            column = columnListIter.next();

            // The check for bigInteger being null is actually motivated by
            // the call to getSpace(). That method will return null for the
            // final column in the section as it does not have a space following
            // it, there is no record in the mark-up and so a null value is
            // returned. To date, I have not seen this happen with the call to
            // getW() but it might.
            bigInteger = column.getW();
            if (bigInteger != null) {
                System.out.println("Column width: " +
                        this.convertSize(bigInteger.longValue(), formatter));
            }
            bigInteger = column.getSpace();
            if (bigInteger != null) {
                System.out.println("Inter-column space: " +
                        this.convertSize(bigInteger.longValue(), formatter));
            }
        }
    }

    /**
     * Microsoft use a standard unit - 1/20th of a point  I think - to
     * record the dimensions of various features of the document within the
     * xml markup. This method converts them into more familiar units - inches
     * and centimetres.
     *
     * @param longValue A primitive long that stores the value to be converted.
     * @param formatter An instance of the DecimalFormat class that is used to
     *                  determine how the value should appear to the user.
     * @return An instance of the String class that encpsulates a message
     * containing the converted value.
     */
    private String convertSize(long longValue, DecimalFormat formatter) {
        return (this.convertSize((double) longValue, formatter));
    }

    /**
     * Microsoft use a standrd unit - 1/20th of a point - to record the dimensions
     * of various features of the document - tabs, page size, margins etc -
     * within the xml markup. This method converts them into more familiar
     * units - inches and centimetres.
     *
     * @param doubleValue A primitive double that stores the value to be
     *                    converted.
     * @param formatter   An instance of the DecimalFormat class that is used to
     *                    determine how the value should appear to the user.
     * @return An instance of the String class that encpsulates a message
     * containing the converted value.
     */
    private String convertSize(double doubleValue, DecimalFormat formatter) {
        StringBuffer buffer = new StringBuffer();
        double sizePoints = doubleValue / TO_POINTS_DIVISOR;
        double sizeInches = sizePoints / TO_INCHES_DIVISOR;
        double sizeCM = sizeInches * TO_CM_MULTIPLIER;
        buffer.append(formatter.format(sizeInches));
        buffer.append("inches or ");
        buffer.append(formatter.format(sizeCM));
        buffer.append("cm.");
        return (buffer.toString());
    }

    public static void main(String[] args) {
        if (args.length != 1) {
            System.out.println("Usage new XWPFSectionTest(new String[]{\"filename\"})");
            System.out.println("where the filename parameter is an instance of the");
            System.out.println("String class that encapsulates the path to and name");
            System.out.println("a valid Word (docx) document.");
        } else {
            try {
                new XWPFSectionTest(args[0]);
            } catch (IOException ioEx) {
                System.out.println("Caught an: " + ioEx.getClass().getName());
                System.out.println("Message: " + ioEx.getMessage());
                System.out.println("Stacktrace follows:.....");
                ioEx.printStackTrace(System.out);
            }
        }
    }
}