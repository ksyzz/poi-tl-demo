package com.ksyzz.poi;

import com.deepoove.poi.NiceXWPFDocument;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.template.run.RunTemplate;
import org.apache.poi.POIXMLException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlToken;
import org.apache.xmlbeans.impl.values.XmlAnyTypeImpl;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

/**
 * @author fengqian
 * @since <pre>2019/08/23</pre>
 */
public class MyPictureRenderPolicy extends AbstractRenderPolicy<PictureRenderData> {

    @Override
    protected boolean validate(PictureRenderData data) {
        return (null != data.getData() || null != data.getPath());
    }

    @Override
    public void doRender(RunTemplate runTemplate, PictureRenderData picture, XWPFTemplate template)
            throws Exception {
        XWPFRun run = runTemplate.getRun();
        MyPictureRenderPolicy.Helper.renderPicture(run, picture);
    }

    @Override
    protected void afterRender(RenderContext context) {
        clearPlaceholder(context, false);
    }

    @Override
    protected void doRenderException(RunTemplate runTemplate, PictureRenderData data, Exception e) {
        logger.info("Render picture " + runTemplate + " error: {}", e.getMessage());
        runTemplate.getRun().setText(data.getAltMeta(), 0);
    }

    public static class Helper {
        public static void renderPicture(XWPFRun run, PictureRenderData picture) throws Exception {
            int suggestFileType = suggestFileType(picture.getPath());
            InputStream ins = null == picture.getData() ? new FileInputStream(picture.getPath())
                    : new ByteArrayInputStream(picture.getData());

            String relationId = run.getDocument().addPictureData(ins, suggestFileType);
            long width = Units.toEMU(picture.getWidth());
            long height = Units.toEMU(picture.getHeight());
            CTDrawing drawing = run.getCTR().addNewDrawing();
            String xml = "<wp:anchor allowOverlap=\"0\" layoutInCell=\"1\" locked=\"0\" behindDoc=\"0\" relativeHeight=\"0\" simplePos=\"0\" distR=\"0\" distL=\"0\" distB=\"0\" distT=\"0\" " +
                    " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"" +
                    " xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\"" +
                    " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"  >" +
                    "<wp:simplePos y=\"0\" x=\"0\"/>" +
                    "<wp:positionH relativeFrom=\"column\">" +
                    "<wp:align>center</wp:align>" +
                    "</wp:positionH>" +
                    "<wp:positionV relativeFrom=\"paragraph\">" +
                    "<wp:posOffset>0</wp:posOffset>" +
                    "</wp:positionV>" +
                    "<wp:extent cy=\""+height+"\" cx=\""+width+"\"/>" +
                    "<wp:effectExtent b=\"0\" r=\"0\" t=\"0\" l=\"0\"/>" +
                    "<wp:wrapTopAndBottom/>" +
                    "<wp:docPr descr=\"Picture Alt\" name=\"Picture Hit\" id=\"0\"/>" +
                    "<wp:cNvGraphicFramePr>" +
                    "<a:graphicFrameLocks noChangeAspect=\"true\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" />" +
                    "</wp:cNvGraphicFramePr>" +
                    "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                    "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                    "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                    "<pic:nvPicPr>" +
                    "<pic:cNvPr name=\"Picture Hit\" id=\"1\"/>" +
                    "<pic:cNvPicPr/>" +
                    "</pic:nvPicPr>" +
                    "<pic:blipFill>" +
                    "<a:blip r:embed=\""+relationId+"\"/>" +
                    "<a:stretch>" +
                    "<a:fillRect/>" +
                    "</a:stretch>" +
                    "</pic:blipFill>" +
                    "<pic:spPr>" +
                    "<a:xfrm>" +
                    "<a:off y=\"0\" x=\"0\"/>" +
                    "<a:ext cy=\""+height+"\" cx=\""+width+"\"/>" +
                    "</a:xfrm>" +
                    "<a:prstGeom prst=\"rect\">" +
                    "<a:avLst/>" +
                    "</a:prstGeom>" +
                    "</pic:spPr>" +
                    "</pic:pic>" +
                    "</a:graphicData>" +
                    "</a:graphic>" +
                    "<wp14:sizeRelH relativeFrom=\"margin\">" +
                    "<wp14:pctWidth>0</wp14:pctWidth>" +
                    "</wp14:sizeRelH>" +
                    "<wp14:sizeRelV relativeFrom=\"margin\">" +
                    "<wp14:pctHeight>0</wp14:pctHeight>" +
                    "</wp14:sizeRelV>" +
                    "</wp:anchor>";
            drawing.set(XmlToken.Factory.parse(xml, DEFAULT_XML_OPTIONS));
            CTPicture pic = getCTPictures(drawing).get(0);
            XWPFPicture xwpfPicture = new XWPFPicture(pic, run);
            run.getEmbeddedPictures().add(xwpfPicture);
        }


        public static List<CTPicture> getCTPictures(XmlObject o) {
            List<CTPicture> pictures = new ArrayList<>();
            XmlObject[] picts = o.selectPath("declare namespace pic='"
                    + CTPicture.type.getName().getNamespaceURI() + "' .//pic:pic");
            for (XmlObject pict : picts) {
                if (pict instanceof XmlAnyTypeImpl) {
                    // Pesky XmlBeans bug - see Bugzilla #49934
                    try {
                        pict = CTPicture.Factory.parse(pict.toString(),
                                DEFAULT_XML_OPTIONS);
                    } catch (XmlException e) {
                        throw new POIXMLException(e);
                    }
                }
                if (pict instanceof CTPicture) {
                    pictures.add((CTPicture) pict);
                }
            }
            return pictures;
        }


        public static int suggestFileType(String imgFile) {
            int format = 0;
            if (imgFile.endsWith(".emf")) {
                format = XWPFDocument.PICTURE_TYPE_EMF;
            } else if (imgFile.endsWith(".wmf")) {
                format = XWPFDocument.PICTURE_TYPE_WMF;
            } else if (imgFile.endsWith(".pict")) {
                format = XWPFDocument.PICTURE_TYPE_PICT;
            } else if (imgFile.endsWith(".jpeg") || imgFile.endsWith(".jpg")) {
                format = XWPFDocument.PICTURE_TYPE_JPEG;
            } else if (imgFile.endsWith(".png")) {
                format = XWPFDocument.PICTURE_TYPE_PNG;
            } else if (imgFile.endsWith(".dib")) {
                format = XWPFDocument.PICTURE_TYPE_DIB;
            } else if (imgFile.endsWith(".gif")) {
                format = XWPFDocument.PICTURE_TYPE_GIF;
            } else if (imgFile.endsWith(".tiff")) {
                format = XWPFDocument.PICTURE_TYPE_TIFF;
            } else if (imgFile.endsWith(".eps")) {
                format = XWPFDocument.PICTURE_TYPE_EPS;
            } else if (imgFile.endsWith(".bmp")) {
                format = XWPFDocument.PICTURE_TYPE_BMP;
            } else if (imgFile.endsWith(".wpg")) {
                format = XWPFDocument.PICTURE_TYPE_WPG;
            } else {
                throw new RenderException(
                        "Unsupported picture: " + imgFile + ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");
            }
            return format;
        }

    }
}
