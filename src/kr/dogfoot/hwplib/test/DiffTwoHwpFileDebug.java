package kr.dogfoot.hwplib.test;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.reader.HWPReader;

public class DiffTwoHwpFileDebug {
    public static void main(String[] args) throws Exception {
        HWPFile hwpFile1 = HWPReader.fromFile("C:\\Users\\DELL\\Downloads\\t1.hwp");
        HWPFile hwpFile2 = HWPReader.fromFile("H:\\project\\testHWP\\sample_hwp\\test-simple-table.hwp");
        System.out.println(hwpFile1);
    }
}
