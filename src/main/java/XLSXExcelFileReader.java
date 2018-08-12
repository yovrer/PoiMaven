import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;


import java.io.BufferedInputStream;
import java.io.InputStream;
import java.util.*;

/**
 * @author  Suhaib Jamil Abu Shawish
 */
public class XLSXExcelFileReader {

    private String file;

    public XLSXExcelFileReader(String file) {
        this.file = file;
    }

    public List<List<String>> read() {
        try {
            return processFirstSheet(file);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private List<String[]> readSheet(Sheet sheet) {
        List<String[]> res = new LinkedList<>();
        Iterator<Row> rowIterator = sheet.rowIterator();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            int cellsNumber = row.getLastCellNum();
            String[] cellsValues = new String[cellsNumber];

            Iterator<Cell> cellIterator = row.cellIterator();
            int cellIndex = 0;

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                cellsValues[cellIndex++] = cell.getStringCellValue();
            }

            res.add(cellsValues);
        }
        return res;
    }

    public String getFile() {
        return file;
    }

    public void setFile(String file) {
        this.file = file;
    }

    private List<List<String>> processFirstSheet(String filename) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ);
        XSSFReader r = new XSSFReader(pkg);
        SharedStringsTable sst = r.getSharedStringsTable();

        SheetHandler handler = new SheetHandler(sst);
        XMLReader parser = fetchSheetParser(handler);
        Iterator<InputStream> sheetIterator = r.getSheetsData();

        if (!sheetIterator.hasNext()) {
            return Collections.emptyList();
        }

        InputStream sheetInputStream = sheetIterator.next();
        BufferedInputStream bisSheet = new BufferedInputStream(sheetInputStream);
        InputSource sheetSource = new InputSource(bisSheet);
        parser.parse(sheetSource);
        List<List<String>> res = handler.getRowCache();
        bisSheet.close();
        return res;
    }

    public XMLReader fetchSheetParser(ContentHandler handler) throws SAXException {
        XMLReader parser =
                XMLReaderFactory.createXMLReader(
                        "org.apache.xerces.parsers.SAXParser"
                );
        parser.setContentHandler(handler);
        return parser;
    }

    /**
     * See org.xml.sax.helpers.DefaultHandler javadocs
     */
    private static class SheetHandler extends DefaultHandler {

        private static final String ROW_EVENT = "row";
        private static final String CELL_EVENT = "c";

        private boolean foundCell = false;

        private boolean foundCellValue = false;

        private SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;

        private List<String> cellCache = new LinkedList<>();
        private List<List<String>> rowCache = new LinkedList<>();

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            // c => cell
            if (CELL_EVENT.equals(name)) {

                if (foundCell && !foundCellValue) {
                    cellCache.add(null);
                }
                foundCellValue = false;
                foundCell = true;
                String cellType = attributes.getValue("t");
                if (cellType != null && cellType.equals("s")) {
                    nextIsString = true;
                } else {
                    nextIsString = false;
                }
            }

            // Clear contents cache
            lastContents = "";
        }

        public void endElement(String uri, String localName, String name) throws SAXException {
            // Process the last contents as required.
            // Do now, as characters() may be called more than once
            if (nextIsString) {
                int index = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(index)).toString();
                nextIsString = false;
            }

            // store the content of cellCache to  rowCache after the row is end.
            if (ROW_EVENT.equals(name)) {
                foundCellValue = false;
                foundCell = false;
                clearLastNullFromList();
                if (!cellCache.isEmpty()) {
                    rowCache.add(new ArrayList<>(cellCache));
                }
                cellCache.clear();
            }

            // v => contents of a cell
            // Output after we've seen the string contents
            if (name.equals("v")) {
                foundCellValue = true;
                cellCache.add(lastContents);
            }
        }

        public void characters(char[] ch, int start, int length)
                throws SAXException {
            lastContents += new String(ch, start, length);
        }

        public List<List<String>> getRowCache() {
            return rowCache;
        }

        private void clearLastNullFromList() {
            while (!cellCache.isEmpty() && cellCache.get(cellCache.size() - 1) == null) {
                cellCache.remove(cellCache.size() - 1);
            }
        }
    }

    public static void main(String[] args) {
        XLSXExcelFileReader XLSXExcelFileReader = new XLSXExcelFileReader("C:\\Users\\rn-sshawish\\Downloads\\7MB.xlsx");
        XLSXExcelFileReader.read().forEach((values) -> {
            for (String value : values) {
                System.out.print(value + " ");
            }
            System.out.println();
        });
    }
}