package kr.dogfoot.hwplib.test;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.reader.HWPReader;

public class DiffTwoHwpFileDebug {
    public static void main(String[] args) throws Exception {
        HWPFile hwpFile1 = HWPReader.fromFile("E:\\d6.hwp");
        HWPFile hwpFile2 = HWPReader.fromFile("E:\\d7.hwp");
        System.out.println(hwpFile1);
    }
}
