package kr.dogfoot.hwplib.test;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.control.ControlTable;
import kr.dogfoot.hwplib.object.bodytext.control.ControlType;
import kr.dogfoot.hwplib.object.bodytext.control.table.Cell;
import kr.dogfoot.hwplib.object.bodytext.control.table.Row;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.reader.HWPReader;
import kr.dogfoot.hwplib.writer.HWPWriter;

import static kr.dogfoot.hwplib.test.TestInsertImage2Table.mmToHwp;

public class TestSimpleTableAdd {
    public static void main(String[] args) throws Exception {
        String filename = "sample_hwp\\test-blank.hwp";

        HWPFile hwpFile = HWPReader.fromFile(filename);
        if (hwpFile != null) {
            Section firstSection = hwpFile.getBodyText().getSectionList().get(0);
            Paragraph firstParagraph1 = firstSection.getParagraph(0);
            firstParagraph1.getText().addExtendCharForTable();
            ControlTable table = (ControlTable) firstParagraph1.addNewControl(ControlType.Table);
            for (int i = 0; i < 44; i++) {
                Row row = table.addNewRow();
                for (int j = 0; j < 4; j++) {
                    Cell cell = row.addNewCell();
                    cell.setWidthAndHeight(mmToHwp(3), mmToHwp(3));
                    if(i%2==0){
                        cell.createTextControl("t "+i+" -- "+j);
                    }else{
                        cell.createTextControl("");
                    }
                }
            }
            table.computeColumnCount();
            String writePath = "sample_hwp\\test-simple-table.hwp";
            HWPWriter.toFile(hwpFile, writePath);
            System.out.println("done");
        }
    }
}
