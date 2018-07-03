package kr.dogfoot.hwplib.test;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.control.ControlTable;
import kr.dogfoot.hwplib.object.bodytext.control.ControlType;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.CtrlHeaderGso;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.*;
import kr.dogfoot.hwplib.object.bodytext.control.gso.ControlPicture;
import kr.dogfoot.hwplib.object.bodytext.control.gso.GsoControlType;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.ShapeComponentNormal;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.lineinfo.*;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.shadowinfo.ShadowInfo;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.shadowinfo.ShadowType;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponenteach.ShapeComponentPicture;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponenteach.polygon.PositionXY;
import kr.dogfoot.hwplib.object.bodytext.control.table.Cell;
import kr.dogfoot.hwplib.object.bodytext.control.table.DivideAtPageBoundary;
import kr.dogfoot.hwplib.object.bodytext.control.table.Row;
import kr.dogfoot.hwplib.object.bodytext.control.table.Table;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.charshape.ParaCharShape;
import kr.dogfoot.hwplib.object.bodytext.paragraph.header.ParaHeader;
import kr.dogfoot.hwplib.object.bodytext.paragraph.lineseg.LineSegItem;
import kr.dogfoot.hwplib.object.bodytext.paragraph.lineseg.ParaLineSeg;
import kr.dogfoot.hwplib.object.docinfo.BinData;
import kr.dogfoot.hwplib.object.docinfo.bindata.BinDataCompress;
import kr.dogfoot.hwplib.object.docinfo.bindata.BinDataState;
import kr.dogfoot.hwplib.object.docinfo.bindata.BinDataType;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.PictureEffect;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.PictureInfo;
import kr.dogfoot.hwplib.object.etc.Color4Byte;
import kr.dogfoot.hwplib.reader.HWPReader;
import kr.dogfoot.hwplib.writer.HWPWriter;
import org.apache.poi.util.IOUtils;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class TestInsertImageToTable {
	public static void main(String[] args) throws Exception {
		String filename = "sample_hwp\\test-blank.hwp";

		HWPFile hwpFile = HWPReader.fromFile(filename);
		if (hwpFile != null) {

			Section firstSection = hwpFile.getBodyText().getSectionList().get(0);
			Paragraph firstParagraph = firstSection.getParagraph(0);

			int streamIndex = hwpFile.getBinData().getEmbeddedBinaryDataList().size() + 1;
			String streamName = getStreamName(streamIndex, "jpg");
			byte[] fileBinary = loadFile(); //load image
			Rectangle shapePosition = new Rectangle(0, 0, 30, 30);
			hwpFile.getBinData().addNewEmbeddedBinaryData(streamName, fileBinary);
			int binDataID = addBinDataInDocInfo(hwpFile, streamIndex);


			firstParagraph.getText().addExtendCharForTable();
			ControlTable table = (ControlTable) firstParagraph.addNewControl(ControlType.Table);
			Row row = table.addNewRow();
			Cell cell = row.addNewCell();
			cell.setWidthAndHeight(fromMM(shapePosition.width),fromMM(shapePosition.height));

			Paragraph paragraph = cell.getParagraphList().addNewParagraph();
			ParaHeader ph = paragraph.getHeader();
			ph.setLastInList(true);
			ph.setParaShapeId(1);
			ph.setStyleId((short) 1);;
			ph.getDivideSort().setDivideSection(false);
			ph.getDivideSort().setDivideMultiColumn(false);
			ph.getDivideSort().setDividePage(false);
			ph.getDivideSort().setDivideColumn(false);
			ph.setCharShapeCount(1);
			ph.setRangeTagCount(0);
			ph.setLineAlignCount(1);
			ph.setInstanceID(0);
			ph.setIsMergedByTrack(0);

			paragraph.createText().addExtendCharForGSO();
			ParaLineSeg pls = paragraph.createLineSeg();
			LineSegItem lsi = pls.addNewLineSegItem();

			lsi.setTextStartPositon(0);
			lsi.setLineVerticalPosition(0);
			lsi.setLineHeight(ptToLineHeight(10.0));
			lsi.setTextPartHeight(ptToLineHeight(10.0));
			lsi.setDistanceBaseLineToLineVerticalPosition(ptToLineHeight(10.0 * 0.85));
			lsi.setLineSpace(ptToLineHeight(3.0));
			lsi.setStartPositionFromColumn(0);
			lsi.setSegmentWidth((int) fromMM(50.0));
			lsi.getTag().setFirstSegmentAtLine(true);
			lsi.getTag().setLastSegmentAtLine(true);

			ParaCharShape charShape = paragraph.createCharShape();
			charShape.addParaCharShape(0,1);
//
			ControlPicture controlPicture = (ControlPicture) paragraph.addNewGsoControl(GsoControlType.Picture);
			controlPicture.setBinDataId(binDataID);
			controlPicture.setPostion(shapePosition.width,shapePosition.height);
			controlPicture.setWidthAndHeight(shapePosition);

			table.computeColumnCount();


			String writePath = "sample_hwp\\test-write-image-table.hwp";
			HWPWriter.toFile(hwpFile, writePath);
			System.out.println("done");
		}
	}

	private static int ptToLineHeight(double pt) {
		return (int) (pt * 100.0f);
	}

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

	private static String getStreamName(int streamIndex, String imageFileExt) {
		return "Bin" + String.format("%04X", streamIndex) + "." + imageFileExt;
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

	private static int fromMM(int mm) {
		if (mm == 0) {
			return 1;
		}

		return (int) ((double) mm * 72000.0f / 254.0f + 0.5f);
	}
	private static int fromMM(double mm) {
		if (mm == 0) {
			return 1;
		}

		return (int) (mm * 72000.0f / 254.0f + 0.5f);
	}
}
