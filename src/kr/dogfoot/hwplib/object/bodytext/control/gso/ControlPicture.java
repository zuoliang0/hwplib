package kr.dogfoot.hwplib.object.bodytext.control.gso;

import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.CtrlHeaderGso;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.*;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.ShapeComponent;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponenteach.ShapeComponentPicture;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponenteach.polygon.PositionXY;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.PictureEffect;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.PictureInfo;

import java.awt.*;

/**
 * 그림 개체 컨트롤
 * 
 * @author neolord
 */
public class ControlPicture extends GsoControl {
	/**
	 * 그림 개체 속성
	 */
	private ShapeComponentPicture shapeComponentPicture;

	/**
	 * 생성자
	 */
	public ControlPicture() {
		this(new CtrlHeaderGso());
	}

    /**
     * 생성자
     *
     * @param header 그리기 개체를 위한 컨트롤 헤더
     */
    public ControlPicture(CtrlHeaderGso header) {
        super(header);
        setGsoId(GsoControlType.Picture.getId());

        shapeComponentPicture = new ShapeComponentPicture();

        GsoHeaderProperty prop = header.getProperty();
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

    }

    public void setPostion(int width, int height) {
        PositionXY leftTop = shapeComponentPicture.getLeftTop();
        PositionXY leftBottom = shapeComponentPicture.getLeftBottom();
        PositionXY rightBottom = shapeComponentPicture.getRightBottom();
        PositionXY rightTop = shapeComponentPicture.getRightTop();

        leftTop.setX(0);
        leftTop.setY(0);
        rightTop.setX(fromMM(width));
        rightTop.setY(0);
        leftBottom.setX(0);
        leftBottom.setY(fromMM(height));
        rightBottom.setX(fromMM(width));
        rightBottom.setY(fromMM(height));
        ShapeComponent shapeComponent = getShapeComponent();
        shapeComponent.setLocalFileVersion(1);
        shapeComponent.setWidthAtCreate(fromMM(width));
        shapeComponent.setHeightAtCreate(fromMM(height));
        shapeComponent.setMatrixsNormal();
        shapeComponent.setWidthAtCurrent(fromMM(width));
        shapeComponent.setHeightAtCurrent(fromMM(height));

    }

    public void setBinDataId(int id) {
        PictureInfo pictureInfo = getShapeComponentPicture().getPictureInfo();
        pictureInfo.setBinItemID(id);
        pictureInfo.setContrast((byte) 0);
        pictureInfo.setBrightness((byte) 0);//对比度
        pictureInfo.setEffect(PictureEffect.RealPicture);
    }

    public void setWidthAndHeight(Rectangle shapePosition) {
        getHeader().setyOffset(fromMM(shapePosition.y));
        getHeader().setxOffset(fromMM(shapePosition.x));
        getHeader().setWidth(fromMM(shapePosition.width));
        getHeader().setHeight(fromMM(shapePosition.height));
        shapeComponentPicture.setImageWidth(fromMM(shapePosition.width));
        shapeComponentPicture.setImageHeight(fromMM(shapePosition.height));
    }

    private int fromMM(int mm) {
        if (mm == 0) {
            return 1;
        }

        return (int) ((double) mm * 72000.0f / 254.0f + 0.5f);
    }

    /**
     * 그림 개체의 속성 객체를 반환한다.
     *
     * @return 그림 개체의 속성 객체
     */
    public ShapeComponentPicture getShapeComponentPicture() {
        return shapeComponentPicture;
    }

}
