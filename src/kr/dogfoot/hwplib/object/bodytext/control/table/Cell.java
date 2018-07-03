package kr.dogfoot.hwplib.object.bodytext.control.table;

import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.sectiondefine.TextDirection;
import kr.dogfoot.hwplib.object.bodytext.control.gso.textbox.LineChange;
import kr.dogfoot.hwplib.object.bodytext.control.gso.textbox.TextVerticalAlignment;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.ParagraphList;
import kr.dogfoot.hwplib.object.bodytext.paragraph.charshape.ParaCharShape;
import kr.dogfoot.hwplib.object.bodytext.paragraph.header.ParaHeader;
import kr.dogfoot.hwplib.object.bodytext.paragraph.lineseg.LineSegItem;
import kr.dogfoot.hwplib.object.bodytext.paragraph.lineseg.ParaLineSeg;

/**
 * 표의 셀을 나타내는 객체
 *
 * @author neolord
 */
public class Cell {
    /**
     * 문단 리스트 헤더
     */
    private ListHeaderForCell listHeader;
    /**
     * 문단 리스트
     */
    private ParagraphList paragraphList;
    private int colIndex;
    private int rowIndex;

    /**
     * 생성자
     */
    public Cell() {
        this(0, 0);
    }

    public Cell(int rowIndex, int colIndex) {
        this.colIndex = colIndex;
        this.rowIndex = rowIndex;
        listHeader = new ListHeaderForCell();
        paragraphList = new ParagraphList();
        listHeader.setParaCount(1);
        listHeader.getProperty().setTextDirection(TextDirection.Horizontal);
        listHeader.getProperty().setLineChange(LineChange.Normal);
        listHeader.getProperty().setTextVerticalAlignment(TextVerticalAlignment.Center);
        listHeader.getProperty().setProtectCell(false);
        listHeader.getProperty().setEditableAtFormMode(false);
        listHeader.setColIndex(colIndex);
        listHeader.setRowIndex(rowIndex);
        listHeader.setColSpan(1);
        listHeader.setRowSpan(1);
        listHeader.setFieldName("");
    }

    public void setWidthAndHeight(long width, long height) {
		listHeader.setWidth(width);
		listHeader.setHeight(height);
		listHeader.setTextWidth(width);
    }

    /**
     * 문단 리스트 헤더를 반환한다.
     *
     * @return 문단 리스트 헤더
     */
    public ListHeaderForCell getListHeader() {
        return listHeader;
    }

    public void createPicControl(){
        Paragraph p = getParagraphList().addNewParagraph();
        ParaHeader ph = p.getHeader();
        ph.setLastInList(true);
        // 셀의 문단 모양을 이미 만들어진 문단 모양으로 사용함
        ph.setParaShapeId(1);
        // 셀의 스타일을이미 만들어진 스타일으로 사용함
        ph.setStyleId((short) 1);
        ph.getDivideSort().setDivideSection(false);
        ph.getDivideSort().setDivideMultiColumn(false);
        ph.getDivideSort().setDividePage(false);
        ph.getDivideSort().setDivideColumn(false);
        ph.setCharShapeCount(1);
        ph.setRangeTagCount(0);
        ph.setLineAlignCount(1);
        ph.setInstanceID(0);
        ph.setIsMergedByTrack(0);
        p.createText().addExtendCharForGSO();

        p.createCharShape();

        ParaCharShape pcs = p.getCharShape();
        // 셀의 글자 모양을 이미 만들어진 글자 모양으로 사용함
        pcs.addParaCharShape(0, 1);
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

    private  int ptToLineHeight(double pt) {
        return (int) (pt * 100.0f);
    }
    private static long mmToHwp(double mm) {
        return (long) (mm * 72000.0f / 254.0f + 0.5f);
    }

    /**
     * 문단 리스트를 반환한다.
     *
     * @return 문단 리스트
     */
    public ParagraphList getParagraphList() {
        return paragraphList;
    }
}
