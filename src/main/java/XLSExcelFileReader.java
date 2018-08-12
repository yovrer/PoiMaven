import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

/**
 * @author  Suhaib Jamil Abu Shawish
 */
public class XLSExcelFileReader {

    private InputStream excelFileStream;

    private String libraryId;

    public class XLSEventListener implements HSSFListener {

        private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;

        private FormatTrackingHSSFListener formatListener;

        private SSTRecord sstrec;

        private List<String> cellCache = new LinkedList<>();

        private List<List<String>> rowCache = new LinkedList<>();

        private int lastRowNumber = -1;

        private int lastColumn = 0;

        /**
         * Should we output the formula, or the value it has?
         */
        private boolean outputFormulaValues = true;

        /**
         * This method listens for incoming records and handles them as required.
         *
         * @param record The record that was found while reading.
         */
        @Override
        public void processRecord(Record record) {

            int currentRow = -1;

            int currentColumn = -1;

            String currentValue = null;

            switch (record.getSid()) {
                // the BOFRecord can represent either the beginning of a sheet or the workbook
                case BOFRecord.sid: {
                    BOFRecord bof = (BOFRecord) record;
                    clearLastNullFromList();
                    if (!cellCache.isEmpty()) {
                        rowCache.add(new ArrayList<>(cellCache));
                    }
                    cellCache.clear();
                    lastRowNumber = currentRow;
                    break;
                }
                case NumberRecord.sid: {
                    NumberRecord numrec = (NumberRecord) record;
                    currentValue = String.format("%.2f", numrec.getValue());
                    currentRow = numrec.getRow();
                    currentColumn = numrec.getColumn();
                    break;
                }
                // SSTRecords store a array of unique strings used in Excel.
                case SSTRecord.sid:
                    sstrec = (SSTRecord) record;
                    break;
                case LabelSSTRecord.sid: {
                    LabelSSTRecord lrec = (LabelSSTRecord) record;
                    currentRow = lrec.getRow();
                    currentValue = sstrec.getString(lrec.getSSTIndex()).toString().trim();
                    currentColumn = lrec.getColumn();
                    break;
                }
            }

            if (currentRow != -1 && lastRowNumber != -1 && currentRow != lastRowNumber) {
                clearLastNullFromList();
                if (!cellCache.isEmpty()) {
                    rowCache.add(new ArrayList<>(cellCache));
                }
                cellCache.clear();
                lastColumn = 0;
                lastRowNumber = currentRow;
            } else if (lastRowNumber == -1 && currentRow != -1) {
                lastRowNumber = currentRow;
            }

            if (currentValue != null && currentValue.length() != 0) {
                fillEmtyCell(lastColumn, currentColumn);
                lastColumn = currentColumn;
                cellCache.add(currentValue);
            }
        }

        private void clearLastNullFromList() {
            while (!cellCache.isEmpty() && cellCache.get(cellCache.size() - 1) == null) {
                cellCache.remove(cellCache.size() - 1);
            }
        }

        private void fillEmtyCell(int start, int end) {
            for (int i = start; i < end - 1; i++) {
                cellCache.add(null);
            }
        }

        public List<List<String>> getRowCache() {
            clearLastNullFromList();
            if (!cellCache.isEmpty()) {
                rowCache.add(new ArrayList<>(cellCache));
            }
            cellCache.clear();
            return rowCache;
        }

    }

    public void startProcessing() {
        // create a new file input stream with the input file specified
        // at the command line
        // create a new org.apache.poi.poifs.filesystem.Filesystem
        try {
            POIFSFileSystem poifs = new POIFSFileSystem(excelFileStream);
            // get the Workbook (excel part) stream in a InputStream
            try (InputStream din = poifs.createDocumentInputStream("Workbook")) {
                // construct out HSSFRequest object
                HSSFRequest req = new HSSFRequest();
                // lazy listen for ALL records with the listener shown above

                XLSEventListener XLSEventListener = new XLSEventListener();

                req.addListenerForAllRecords(XLSEventListener);
                // create our event factory
                HSSFEventFactory factory = new HSSFEventFactory();
                // process our events based on the document input stream
                factory.processEvents(req, din);

                XLSEventListener.getRowCache().forEach((values) -> {
                    for (String value : values) {
                        System.out.print(value + " ");
                    }
                    System.out.println();
                });
            }
        } catch (Exception s) {
            s.printStackTrace();
        }
    }
}