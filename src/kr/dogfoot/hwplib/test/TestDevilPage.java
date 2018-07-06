package kr.dogfoot.hwplib.test;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.reader.HWPReader;

public class TestDevilPage {
    public static void main(String[] args) throws Exception {
        String filename = "sample_hwp\\test-blank.hwp";

        HWPFile hwpFile = HWPReader.fromFile(filename);
        if (hwpFile != null) {
            Section f = hwpFile.getBodyText().getSectionList().get(0);
            Paragraph paragraph = f.getParagraph(0);
            Paragraph secondParagraph = f.addNewParagraph();
            secondParagraph.getHeader().getDivideSort().setDivideColumn(true);
//            secondParagraph.getHeader().getDivideSort().getValue();
//            secondParagraph.getHeader().getDivideSort().setDividePage(true);
        }


    }
}
