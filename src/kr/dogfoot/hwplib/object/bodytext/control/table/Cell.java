package kr.dogfoot.hwplib.object.bodytext.control.table;

import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.sectiondefine.TextDirection;
import kr.dogfoot.hwplib.object.bodytext.control.gso.textbox.LineChange;
import kr.dogfoot.hwplib.object.bodytext.control.gso.textbox.TextVerticalAlignment;
import kr.dogfoot.hwplib.object.bodytext.paragraph.ParagraphList;

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

    public void setWidthAndHeight(int width, int height) {
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

    /**
     * 문단 리스트를 반환한다.
     *
     * @return 문단 리스트
     */
    public ParagraphList getParagraphList() {
        return paragraphList;
    }
}
