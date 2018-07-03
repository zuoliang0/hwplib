package kr.dogfoot.hwplib.object.bodytext.control;

import java.util.ArrayList;

import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.CtrlHeader;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.CtrlHeaderGso;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.*;
import kr.dogfoot.hwplib.object.bodytext.control.gso.caption.Caption;
import kr.dogfoot.hwplib.object.bodytext.control.table.DivideAtPageBoundary;
import kr.dogfoot.hwplib.object.bodytext.control.table.Row;
import kr.dogfoot.hwplib.object.bodytext.control.table.Table;

/**
 * 표 컨트롤
 * 
 * @author neolord
 */
public class ControlTable extends Control {
	/**
	 * 캡션
	 */
	private Caption caption;
	/**
	 * 표 정보
	 */
	private Table table;
	/**
	 * 행 리스트
	 */
	private ArrayList<Row> rowList;

	/**
	 * 생성자
	 */
	public ControlTable() {
		this(new CtrlHeaderGso(ControlType.Table));
		
		caption = null;
		table = new Table();
		rowList = new ArrayList<Row>();
	}

	/**
	 * 생성자
	 * 
	 * @param header
	 *            컨트롤 헤더
	 */
	public ControlTable(CtrlHeader ctrlHeader) {
		super(ctrlHeader);
		CtrlHeaderGso header=(CtrlHeaderGso) ctrlHeader;
		caption = null;
		table = new Table();
		rowList = new ArrayList<Row>();
		header.getProperty().setLikeWord(false);
		header.getProperty().setApplyLineSpace(false);
		header.getProperty().setVertRelTo(VertRelTo.Para);
		header.getProperty().setVertRelativeArrange(RelativeArrange.TopOrLeft);
		header.getProperty().setHorzRelTo(HorzRelTo.Para);
		header.getProperty().setHorzRelativeArrange(RelativeArrange.TopOrLeft);
		header.getProperty().setVertRelToParaLimit(false);
		header.getProperty().setAllowOverlap(false);
		header.getProperty().setWidthCriterion(WidthCriterion.Absolute);
		header.getProperty().setHeightCriterion(HeightCriterion.Absolute);
		header.getProperty().setProtectSize(false);
		header.getProperty().setTextFlowMethod(TextFlowMethod.Tight);
		header.getProperty().setTextHorzArrange(TextHorzArrange.BothSides);
		header.getProperty().setObjectNumberSort(ObjectNumberSort.Table);
		header.setxOffset(fromMM(0));
		header.setyOffset(fromMM(0));
		header.setOutterMarginLeft(0);
		header.setOutterMarginRight(0);
		header.setOutterMarginTop(0);
		header.setOutterMarginBottom(0);
		table.getProperty().setDivideAtPageBoundary(DivideAtPageBoundary.DivideByCell);
		table.getProperty().setAutoRepeatTitleRow(false);
		table.setCellSpacing(0);
		table.setLeftInnerMargin(0);
		table.setRightInnerMargin(0);
		table.setTopInnerMargin(0);
		table.setBottomInnerMargin(0);
		table.setBorderFillId(0);
		table.getCellCountOfRowList().add(1);
	}
	private   int fromMM(int mm) {
		if (mm == 0) {
			return 1;
		}

		return (int) ((double) mm * 72000.0f / 254.0f + 0.5f);
	}
	/**
	 * 그리기 객체 용 컨트롤 헤더를 반환한다.
	 * 
	 * @return 그리기 객체 용 컨트롤 헤더
	 */
	public CtrlHeaderGso getHeader() {
		return (CtrlHeaderGso) header;
	}
	
	/**
	 * 캡션 객체를 생성한다.
	 */
	public void createCaption() {
		caption = new Caption();
	}

	/**
	 * 캡션 객체를 삭제한다.
	 */
	public void deleteCaption() {
		caption = null;
	}

	/**
	 * 캡션 객체를 반환한다.
	 * 
	 * @return 캡션 객체
	 */
	public Caption getCaption() {
		return caption;
	}

	/**
	 * 표 정보 객체를 반환한다.
	 * 
	 * @return 표 정보 객체
	 */
	public Table getTable() {
		return table;
	}

	/**
	 * 새로운 행 객체를 생성하고 리스트에 추가한다.
	 * 
	 * @return 새로 생성된 행 객체
	 */
	public Row addNewRow() {
		Row r = new Row(rowList.size());
		table.setRowCount(table.getRowCount()+1);
		rowList.add(r);
		return r;
	}

	/**
	 * 添加完成列之后计算第一列数量为列数量
	 */
	public void computeColumnCount(){
		if (!this.rowList.isEmpty()) {
			table.setColumnCount(this.rowList.get(0).getCellList().size());
		}
	}
	/**
	 * 행 리스트를 반환한다.
	 * 
	 * @return 행 리스트
	 */
	public ArrayList<Row> getRowList() {
		return rowList;
	}
}
