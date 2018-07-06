package kr.dogfoot.hwplib.test;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.control.ControlTable;
import kr.dogfoot.hwplib.object.bodytext.control.ControlType;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.CtrlHeaderGso;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.*;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.sectiondefine.TextDirection;
import kr.dogfoot.hwplib.object.bodytext.control.gso.ControlPicture;
import kr.dogfoot.hwplib.object.bodytext.control.gso.GsoControlType;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.ShapeComponentNormal;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.lineinfo.*;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.shadowinfo.ShadowInfo;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.shadowinfo.ShadowType;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponenteach.ShapeComponentRectangle;
import kr.dogfoot.hwplib.object.bodytext.control.gso.textbox.LineChange;
import kr.dogfoot.hwplib.object.bodytext.control.gso.textbox.TextVerticalAlignment;
import kr.dogfoot.hwplib.object.bodytext.control.table.*;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.charshape.ParaCharShape;
import kr.dogfoot.hwplib.object.bodytext.paragraph.header.ParaHeader;
import kr.dogfoot.hwplib.object.bodytext.paragraph.lineseg.LineSegItem;
import kr.dogfoot.hwplib.object.bodytext.paragraph.lineseg.ParaLineSeg;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.ParaText;
import kr.dogfoot.hwplib.object.docinfo.BinData;
import kr.dogfoot.hwplib.object.docinfo.BorderFill;
import kr.dogfoot.hwplib.object.docinfo.bindata.BinDataCompress;
import kr.dogfoot.hwplib.object.docinfo.bindata.BinDataState;
import kr.dogfoot.hwplib.object.docinfo.bindata.BinDataType;
import kr.dogfoot.hwplib.object.docinfo.borderfill.BackSlashDiagonalShape;
import kr.dogfoot.hwplib.object.docinfo.borderfill.BorderThickness;
import kr.dogfoot.hwplib.object.docinfo.borderfill.BorderType;
import kr.dogfoot.hwplib.object.docinfo.borderfill.SlashDiagonalShape;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.*;
import kr.dogfoot.hwplib.reader.HWPReader;
import kr.dogfoot.hwplib.writer.HWPWriter;
import org.apache.poi.util.IOUtils;

import java.awt.*;
import java.io.*;

public class TestOwenTable {
    private HWPFile hwpFile;
    private ControlTable table;
    private Row row;
    private Cell cell;
    private int borderFillIDForCell;

    public static void main(String[] args) throws Exception {
        String filename = "sample_hwp\\test-blank.hwp";

        HWPFile hwpFile = HWPReader.fromFile(filename);
        if (hwpFile != null) {

            Section firstSection = hwpFile.getBodyText().getSectionList().get(0);
            Paragraph firstParagraph = firstSection.getParagraph(0);

            // 문단에서 표 컨트롤의 위치를 표현하기 위한 확장 문자를 넣는다. 要在段落中扩展表格控件的位置，可以通过插入一个字符来停止它。
            firstParagraph.getText().addExtendCharForTable();

            // 문단에 표 컨트롤 추가한다.
            ControlTable table = (ControlTable) firstParagraph.addNewControl(ControlType.Table);
//			createTableControlAtFirstParagraph();


//			setCtrlHeaderRecord();
//            int zOrder = 0;
//            CtrlHeaderGso ctrlHeader = table.getHeader();
//            ctrlHeader.getProperty().setLikeWord(false);
//            ctrlHeader.getProperty().setApplyLineSpace(false);
//            ctrlHeader.getProperty().setVertRelTo(VertRelTo.Para);
//            ctrlHeader.getProperty().setVertRelativeArrange(RelativeArrange.TopOrLeft);
//            ctrlHeader.getProperty().setHorzRelTo(HorzRelTo.Para);
//            ctrlHeader.getProperty().setHorzRelativeArrange(RelativeArrange.TopOrLeft);
//            ctrlHeader.getProperty().setVertRelToParaLimit(false);
//            ctrlHeader.getProperty().setAllowOverlap(false);
//            ctrlHeader.getProperty().setWidthCriterion(WidthCriterion.Absolute);
//            ctrlHeader.getProperty().setHeightCriterion(HeightCriterion.Absolute);
//            ctrlHeader.getProperty().setProtectSize(false);
//            ctrlHeader.getProperty().setTextFlowMethod(TextFlowMethod.Tight);
//            ctrlHeader.getProperty().setTextHorzArrange(TextHorzArrange.BothSides);
//            ctrlHeader.getProperty().setObjectNumberSort(ObjectNumberSort.Table);
//            ctrlHeader.setxOffset(mmToHwp(0.0));
//            ctrlHeader.setyOffset(mmToHwp(0.0));
//            ctrlHeader.setWidth(mmToHwp(100.0));
//            ctrlHeader.setHeight(mmToHwp(60.0));
//            ctrlHeader.setzOrder(zOrder++);
//            ctrlHeader.setOutterMarginLeft(0);
//            ctrlHeader.setOutterMarginRight(0);
//            ctrlHeader.setOutterMarginTop(0);
//            ctrlHeader.setOutterMarginBottom(0);

//			setTableRecordFor2By2Cells();
            Table tableRecord = table.getTable();
            tableRecord.getProperty().setDivideAtPageBoundary(DivideAtPageBoundary.DivideByCell);
            tableRecord.getProperty().setAutoRepeatTitleRow(false);
//            tableRecord.setRowCount(10);
//            tableRecord.setColumnCount(10);
            tableRecord.setCellSpacing(0);
            tableRecord.setLeftInnerMargin(0);
            tableRecord.setRightInnerMargin(0);
            tableRecord.setTopInnerMargin(0);
            tableRecord.setBottomInnerMargin(0);
            tableRecord.setBorderFillId(getBorderFillIDForTableOutterLine(hwpFile));
//            tableRecord.getCellCountOfRowList().add(10);
//			add2By2Cell();
            int borderFillIDForCell = 0;
            //addFirstRow();
//            Row row = table.addNewRow();
//            addCell(row,"当前是0行的0列。happy",0,0,borderFillIDForCell);
//            addCell(row,"当前是0行的1列。happy",0,1,borderFillIDForCell);
//            addCell(row,"当前是0行的2列。happy",0,2,borderFillIDForCell);
//            addCell(row,"当前是0行的3列。happy",0,3,borderFillIDForCell);
            for (int i =1;i<=10;i++){
                Row row = table.addNewRow();
                for(int j=1;j<=10;j++){
                    addCell(row,"当前是"+i+"行的"+j+"列。happy",i-1,j-1,borderFillIDForCell);
                }
            }
//            Row row = table.addNewRow();
//            Cell cell = row.addNewCell();
//            cell.setWidthAndHeight(60,60);




            table.computeColumnCount();
//            int streamIndex = hwpFile.getBinData().getEmbeddedBinaryDataList().size() + 1;// 获取当前文档的最新bin文件id
//            String streamName = getStreamName(streamIndex, "jpg");//设置存放到文件中的文件名
//            byte[] fileBinary = loadFile(); //加载图片
//            Rectangle shapePosition = new Rectangle(0, 3, 30, 30);//建一个框用来存放图片
//            hwpFile.getBinData().addNewEmbeddedBinaryData(streamName, fileBinary);//将文件放到文件合集中
//            int binDataID =addBinDataInDocInfo(hwpFile, streamIndex);
//
//            Paragraph paragraph = cell.getParagraphList().addNewParagraph();
//            paragraph.getText().addExtendCharForGSO();
////            ControlRectangle rectangle = (ControlRectangle) firstParagraph.addNewGsoControl(GsoControlType.Rectangle);
//            ControlPicture controlPicture = (ControlPicture) paragraph.addNewGsoControl(GsoControlType.Picture);
//            controlPicture.setWidthAndHeight(shapePosition);
//            controlPicture.setPostion(shapePosition.width,shapePosition.height);
//            controlPicture.setBinDataId(binDataID);

//            setShapeComponent((ShapeComponentNormal) rectangle.getShapeComponent(), shapePosition,binDataID);
//            setShapeComponentRectangle(rectangle.getShapeComponentRectangle(), shapePosition);

            String writePath = "sample_hwp\\test-my-table.hwp";
            HWPWriter.toFile(hwpFile, writePath);
            System.out.println("done");
        }
    }

    private static void setShapeComponentRectangle(ShapeComponentRectangle scr, Rectangle shapePosition) {
        scr.setRoundRate((byte) 0);
        //左上角
        scr.setX1(0);
        scr.setY1(0);
        //右上角
        scr.setX2(fromMM(shapePosition.width));
        scr.setY2(0);
        //右下角
        scr.setX3(fromMM(shapePosition.width));
        scr.setY3(fromMM(shapePosition.height));
        //左下角
        scr.setX4(0);
        scr.setY4(fromMM(shapePosition.height));
    }

    private static void setShapeComponent(ShapeComponentNormal sc, Rectangle shapePosition, int binDataID) {
        sc.setOffsetX(10);
        sc.setOffsetY(10);
        sc.setGroupingCount(0);
        sc.setLocalFileVersion(1);
        sc.setWidthAtCreate(fromMM(shapePosition.width));
        sc.setHeightAtCreate(fromMM(shapePosition.height));
        sc.setWidthAtCurrent(fromMM(shapePosition.width));
        sc.setHeightAtCurrent(fromMM(shapePosition.height));
        sc.setRotateAngle(0);
        sc.setRotateXCenter(fromMM(shapePosition.width / 2));
        sc.setRotateYCenter(fromMM(shapePosition.height / 2));

        sc.createLineInfo();
        LineInfo li = sc.getLineInfo();
        li.getProperty().setLineEndShape(LineEndShape.Flat);
        li.getProperty().setStartArrowShape(LineArrowShape.None);
        li.getProperty().setStartArrowSize(LineArrowSize.MiddleMiddle);
        li.getProperty().setEndArrowShape(LineArrowShape.None);
        li.getProperty().setEndArrowSize(LineArrowSize.MiddleMiddle);
        ;
        li.getProperty().setFillStartArrow(true);
        li.getProperty().setFillEndArrow(true);
        li.getProperty().setLineType(LineType.None);
        li.setOutlineStyle(OutlineStyle.Normal);
        li.setThickness(0);// 厚度
        li.getColor().setValue(0);

        sc.createFillInfo();
        FillInfo fi = sc.getFillInfo();
        fi.getType().setPatternFill(false);
        fi.getType().setImageFill(true);
        fi.getType().setGradientFill(false);
        fi.createImageFill();
        ImageFill imgF = fi.getImageFill();
        imgF.setImageFillType(ImageFillType.FitSize);
        imgF.getPictureInfo().setBrightness((byte) 0);// 亮度
        imgF.getPictureInfo().setContrast((byte) 0);//对比度
        imgF.getPictureInfo().setEffect(PictureEffect.RealPicture);
        imgF.getPictureInfo().setBinItemID(binDataID);

        sc.createShadowInfo();
        ShadowInfo si = sc.getShadowInfo();
        si.setType(ShadowType.None);
        si.getColor().setValue(0xc4c4c4);
        si.setOffsetX(0);
        si.setOffsetY(0);
        si.setTransparnet((short) 0);

        sc.setMatrixsNormal();
    }

    private static void setCtrlHeaderGso(CtrlHeaderGso hdr, Rectangle shapePosition) {
        GsoHeaderProperty prop = hdr.getProperty();
        prop.setLikeWord(false);
        prop.setApplyLineSpace(false);
        prop.setVertRelTo(VertRelTo.Para);
        prop.setVertRelativeArrange(RelativeArrange.TopOrLeft);
        prop.setHorzRelTo(HorzRelTo.Para);
        prop.setHorzRelativeArrange(RelativeArrange.TopOrLeft);
        prop.setVertRelToParaLimit(true);
        prop.setAllowOverlap(true);
        prop.setWidthCriterion(WidthCriterion.Absolute);
        prop.setHeightCriterion(HeightCriterion.Absolute);
        prop.setProtectSize(false);
        prop.setTextFlowMethod(TextFlowMethod.TopAndBottom);
        prop.setTextHorzArrange(TextHorzArrange.BothSides);
        prop.setObjectNumberSort(ObjectNumberSort.Figure);

        hdr.setyOffset(fromMM(shapePosition.y));
        hdr.setxOffset(fromMM(shapePosition.x));
        hdr.setWidth(fromMM(shapePosition.width));
        hdr.setHeight(fromMM(shapePosition.height));
        hdr.setzOrder(0);
        hdr.setOutterMarginLeft(0);
        hdr.setOutterMarginRight(0);
        hdr.setOutterMarginTop(0);
        hdr.setOutterMarginBottom(0);
        hdr.setInstanceId(0x5bb840e1);// 这个有啥用
        hdr.setPreventPageDivide(false);
        hdr.setExplanation(null);
    }

    private static String getStreamName(int streamIndex, String imageFileExt) {
        return "Bin" + String.format("%04X", streamIndex) + "." + imageFileExt;
    }

    private static int fromMM(int mm) {
        if (mm == 0) {
            return 1;
        }

        return (int) ((double) mm * 72000.0f / 254.0f + 0.5f);
    }

    /**
     *
     * 设置一个图片文件的存放关联信息。
     * @param file
     * @param streamIndex
     * @return
     */
    private static int addBinDataInDocInfo(HWPFile file, int streamIndex) {
        BinData bd = new BinData();
        bd.getProperty().setType(BinDataType.Embedding);
        bd.getProperty().setCompress(BinDataCompress.ByStroageDefault);
        bd.getProperty().setState(BinDataState.NotAcceess);
        bd.setBinDataID(streamIndex);
        bd.setExtensionForEmbedding("jpg");
        file.getDocInfo().getBinDataList().add(bd);
        return file.getDocInfo().getBinDataList().size();
    }

    private static byte[] loadFile() throws IOException {
        File file = new File("sample_hwp\\sample.jpg");
        byte[] buffer = new byte[(int) file.length()];
        InputStream ios = null;
        try {
            ios = new FileInputStream(file);
            IOUtils.readFully(ios, buffer);
        } finally {
            try {
                if (ios != null) {
                    ios.close();
                }
            } catch (IOException e) {
            }
        }
        return buffer;
    }

    private static void addCell(Row row, String text, int rownum, int cellnum, int borderFillIDForCell) {
        Cell cell = row.addNewCell();
        setListHeaderForCell(cell, cellnum, rownum, borderFillIDForCell);
        setParagraphForCell(cell, text);
    }


    private int zOrder = 0;


    private static long mmToHwp(double mm) {
        return (long) (mm * 72000.0f / 254.0f + 0.5f);
    }


    private static int getBorderFillIDForTableOutterLine(HWPFile hwpFile) {
        BorderFill bf = hwpFile.getDocInfo().addNewBorderFill();
        bf.getProperty().set3DEffect(false);
        bf.getProperty().setShadowEffect(false);
        bf.getProperty().setSlashDiagonalShape(SlashDiagonalShape.None);
        bf.getProperty().setBackSlashDiagonalShape(BackSlashDiagonalShape.None);
        bf.getLeftBorder().setType(BorderType.None);
        bf.getLeftBorder().setThickness(BorderThickness.MM0_5);
        bf.getLeftBorder().getColor().setValue(0x0);
        bf.getRightBorder().setType(BorderType.None);
        bf.getRightBorder().setThickness(BorderThickness.MM0_5);
        bf.getRightBorder().getColor().setValue(0x0);
        bf.getTopBorder().setType(BorderType.None);
        bf.getTopBorder().setThickness(BorderThickness.MM0_5);
        bf.getTopBorder().getColor().setValue(0x0);
        bf.getBottomBorder().setType(BorderType.None);
        bf.getBottomBorder().setThickness(BorderThickness.MM0_5);
        bf.getBottomBorder().getColor().setValue(0x0);
        bf.setDiagonalSort(BorderType.None);
        bf.setDiagonalThickness(BorderThickness.MM0_5);
        ;
        bf.getDiagonalColor().setValue(0x0);

        bf.getFillInfo().getType().setPatternFill(true);
        bf.getFillInfo().createPatternFill();
        PatternFill pf = bf.getFillInfo().getPatternFill();
        pf.setPatternType(PatternType.None);
        pf.getBackColor().setValue(-1);
        pf.getPatternColor().setValue(0);

        return hwpFile.getDocInfo().getBorderFillList().size();
    }


    private static int getBorderFillIDForCell(HWPFile hwpFile) {
        BorderFill bf = hwpFile.getDocInfo().addNewBorderFill();
        bf.getProperty().set3DEffect(false);
        bf.getProperty().setShadowEffect(false);
        bf.getProperty().setSlashDiagonalShape(SlashDiagonalShape.None);
        bf.getProperty().setBackSlashDiagonalShape(BackSlashDiagonalShape.None);
        bf.getLeftBorder().setType(BorderType.Solid);
        bf.getLeftBorder().setThickness(BorderThickness.MM0_5);
        bf.getLeftBorder().getColor().setValue(0x0);
        bf.getRightBorder().setType(BorderType.Solid);
        bf.getRightBorder().setThickness(BorderThickness.MM0_5);
        bf.getRightBorder().getColor().setValue(0x0);
        bf.getTopBorder().setType(BorderType.Solid);
        bf.getTopBorder().setThickness(BorderThickness.MM0_5);
        bf.getTopBorder().getColor().setValue(0x0);
        bf.getBottomBorder().setType(BorderType.Solid);
        bf.getBottomBorder().setThickness(BorderThickness.MM0_5);
        bf.getBottomBorder().getColor().setValue(0x0);
        bf.setDiagonalSort(BorderType.None);
        bf.setDiagonalThickness(BorderThickness.MM0_5);
        ;
        bf.getDiagonalColor().setValue(0x0);

        bf.getFillInfo().getType().setPatternFill(true);
        bf.getFillInfo().createPatternFill();
        PatternFill pf = bf.getFillInfo().getPatternFill();
        pf.setPatternType(PatternType.None);
        pf.getBackColor().setValue(-1);
        pf.getPatternColor().setValue(0);

        return hwpFile.getDocInfo().getBorderFillList().size();
    }


    private static void setListHeaderForCell(Cell cell, int colIndex, int rowIndex, int borderFillIDForCell) {
        ListHeaderForCell lh = cell.getListHeader();
        lh.setParaCount(1);
        lh.getProperty().setTextDirection(TextDirection.Horizontal);
        lh.getProperty().setLineChange(LineChange.Normal);
        lh.getProperty().setTextVerticalAlignment(TextVerticalAlignment.Center);
        lh.getProperty().setProtectCell(false);
        lh.getProperty().setEditableAtFormMode(false);
        lh.setColIndex(colIndex);
        lh.setRowIndex(rowIndex);
        lh.setColSpan(1);
        lh.setRowSpan(1);
        lh.setWidth(mmToHwp(30.0));
        lh.setHeight(mmToHwp(15.0));
        lh.setLeftMargin(0);
        lh.setRightMargin(0);
        lh.setTopMargin(0);
        lh.setBottomMargin(0);
        lh.setBorderFillId(borderFillIDForCell);
        lh.setTextWidth(mmToHwp(50.0));
        lh.setFieldName("");
    }

    private static void setParagraphForCell(Cell cell, String text) {
        Paragraph p = cell.getParagraphList().addNewParagraph();
        setParaHeader(p);
        setParaText(p, text);
        setParaCharShape(p);
        setParaLineSeg(p);
    }

    private static void setParaHeader(Paragraph p) {
        ParaHeader ph = p.getHeader();
        ph.setLastInList(true);
        // 셀의 문단 모양을 이미 만들어진 문단 모양으로 사용함
        ph.setParaShapeId(1);
        // 셀의 스타일을이미 만들어진 스타일으로 사용함
        ph.setStyleId((short) 1);
        ;
        ph.getDivideSort().setDivideSection(false);
        ph.getDivideSort().setDivideMultiColumn(false);
        ph.getDivideSort().setDividePage(false);
        ph.getDivideSort().setDivideColumn(false);
        ph.setCharShapeCount(1);
        ph.setRangeTagCount(0);
        ph.setLineAlignCount(1);
        ph.setInstanceID(0);
        ph.setIsMergedByTrack(0);
    }

    private static void setParaText(Paragraph p, String text2) {
        p.createText();
        ParaText pt = p.getText();
        try {
            pt.addString(text2);
        } catch (UnsupportedEncodingException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    private static void setParaCharShape(Paragraph p) {
        p.createCharShape();

        ParaCharShape pcs = p.getCharShape();
        // 셀의 글자 모양을 이미 만들어진 글자 모양으로 사용함
        pcs.addParaCharShape(0, 1);
    }


    private static void setParaLineSeg(Paragraph p) {
        p.createLineSeg();

        ParaLineSeg pls = p.getLineSeg();
        LineSegItem lsi = pls.addNewLineSegItem();

        lsi.setTextStartPositon(0);
        lsi.setLineVerticalPosition(0);
        lsi.setLineHeight(ptToLineHeight(10.0));
        lsi.setTextPartHeight(ptToLineHeight(10.0));
        lsi.setDistanceBaseLineToLineVerticalPosition(ptToLineHeight(10.0 * 0.85));
        lsi.setLineSpace(ptToLineHeight(3.0));
        lsi.setStartPositionFromColumn(0);
        lsi.setSegmentWidth((int) mmToHwp(50.0));
        lsi.getTag().setFirstSegmentAtLine(true);
        lsi.getTag().setLastSegmentAtLine(true);
    }

    private static int ptToLineHeight(double pt) {
        return (int) (pt * 100.0f);
    }


}
