package ch.interlis.ioxwkf.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.vividsolutions.jts.geom.LineString;
import com.vividsolutions.jts.geom.MultiLineString;
import com.vividsolutions.jts.geom.MultiPoint;
import com.vividsolutions.jts.geom.MultiPolygon;
import com.vividsolutions.jts.geom.Point;

import ch.ehi.basics.settings.Settings;
import ch.interlis.ili2c.generator.XSDGenerator;
import ch.interlis.ili2c.metamodel.LocalAttribute;
import ch.interlis.ili2c.metamodel.NumericType;
import ch.interlis.ili2c.metamodel.NumericalType;
import ch.interlis.ili2c.metamodel.TransferDescription;
import ch.interlis.ili2c.metamodel.Viewable;
import ch.interlis.iom.IomObject;
import ch.interlis.iox.IoxEvent;
import ch.interlis.iox.IoxException;
import ch.interlis.iox.IoxFactoryCollection;
import ch.interlis.iox.IoxWriter;
import ch.interlis.iox.ObjectEvent;
import ch.interlis.iox.StartBasketEvent;
import ch.interlis.iox.StartTransferEvent;

public class ExcelWriter implements IoxWriter {
    private File outputFile;

//    private Schema schema = null;
    private Row headerRow = null;
    private List<ExcelAttributeDescriptor> attrDescs = null;
    private XSSFWorkbook workbook = null;
    private XSSFSheet sheet = null;
    private int rowNum = 0;

    private TransferDescription td = null;
//    private String iliGeomAttrName = null;
    
    // ili types
    private static final String COORD="COORD";
    private static final String MULTICOORD="MULTICOORD";
    private static final String POLYLINE="POLYLINE";
    private static final String MULTIPOLYLINE="MULTIPOLYLINE";
    private static final String MULTISURFACE="MULTISURFACE";

    private Integer srsId = null;
    private Integer defaultSrsId = 2056; // TODO: null

    public ExcelWriter(File file) throws IoxException {
        this(file,null);
        System.setProperty("log4j2.loggerContextFactory","org.apache.logging.log4j.simple.SimpleLoggerContextFactory");
    }

    public ExcelWriter(File file, Settings settings) throws IoxException {
        init(file,settings);
    }

    private void init(File file, Settings settings) throws IoxException {
        //this.outputStream = new FileOutputStream(file);
        this.outputFile = file;
    }

    public void setModel(TransferDescription td) {
        this.td = td;
    }

    public void setAttributeDescriptors(List<ExcelAttributeDescriptor> attrDescs) {
        this.attrDescs = attrDescs;
    }
    
    @Override
    public void write(IoxEvent event) throws IoxException {
        if(event instanceof StartTransferEvent){
            // ignore
        } else if(event instanceof StartBasketEvent) {
        } else if(event instanceof ObjectEvent){
            ObjectEvent obj = (ObjectEvent) event;
            IomObject iomObj = obj.getIomObject();
            String tag = iomObj.getobjecttag();
            
            System.out.println("tag: " + tag);
            
            // Wenn null, dann gibt es noch kein "Schema".
            if(attrDescs == null) {
                attrDescs = new ArrayList<>();
                if(td != null) {
                    Viewable aclass = (Viewable) XSDGenerator.getTagMap(td).get(tag);
                    if (aclass == null) {
                        throw new IoxException("class "+iomObj.getobjecttag()+" not found in model");
                    }
                    Iterator viewableIter = aclass.getAttributes();
                    while(viewableIter.hasNext()) {
                        ExcelAttributeDescriptor attrDesc = new ExcelAttributeDescriptor();

                        Object attrObj = viewableIter.next();
                        //System.out.println(attrObj);

                        if(attrObj instanceof LocalAttribute) {
                            LocalAttribute localAttr = (LocalAttribute)attrObj;
                            String attrName = localAttr.getName();
                            //System.out.println(attrName);
                            attrDesc.setAttributeName(attrName);

                            // TODO Geometriedinger
                            ch.interlis.ili2c.metamodel.Type iliType = localAttr.getDomainResolvingAliases();
                            if (iliType instanceof ch.interlis.ili2c.metamodel.NumericalType) {
                                NumericalType numericalType = (NumericalType)iliType;
                                NumericType numericType = (NumericType)numericalType;
                                int precision = numericType.getMinimum().getAccuracy(); 
                                if (precision > 0) {
                                    attrDesc.setBinding(Double.class);
                                } else {
                                    attrDesc.setBinding(Integer.class);
                                }
                                attrDescs.add(attrDesc);
                            } else {
                                if (localAttr.isDomainBoolean()) {
                                    attrDesc.setBinding(Boolean.class);
                                    attrDescs.add(attrDesc);
                                } else if (localAttr.isDomainIli2Date()) {
                                    attrDesc.setBinding(LocalDate.class);
                                    attrDescs.add(attrDesc);
                                } else if (localAttr.isDomainIli2DateTime()) {
                                    attrDesc.setBinding(LocalDateTime.class);
                                    attrDescs.add(attrDesc);
                                } else if (localAttr.isDomainIli2Time()) {
                                    attrDesc.setBinding(LocalTime.class);
                                    attrDescs.add(attrDesc);
                                } else {
                                    attrDesc.setBinding(String.class);
                                    attrDescs.add(attrDesc);
                                }
                            }
                        }
                    }
                } else {
                    for(int u=0;u<iomObj.getattrcount();u++) {
                        String attrName = iomObj.getattrname(u);
                        //create the builder
                        ExcelAttributeDescriptor attrDesc = new ExcelAttributeDescriptor();

                        // Es wurde weder ein Modell gesetzt noch wurde das Schema
                        // mittel setAttrDescs definiert. -> Es wird aus dem ersten IomObject
                        // das Zielschema möglichst gut definiert.
                        // Nachteile:
                        // - Geometrie aus Struktur eruieren ... siehe Kommentar wegen anderen Strukturen. Kann eventuell abgefedert werden.
                        // - Wenn das erste Element fehlende Attribute hat (also NULL-Werte) gehen diese Attribute bei der Schemadefinition
                        // verloren.

                        // Ist das nicht relativ heikel?
                        // Funktioniert mit mehr, wenn es andere Strukturen gibt, oder? Wegen getattrvaluecount?
                        // TODO: testen
                        if (iomObj.getattrvaluecount(attrName)>0 && iomObj.getattrobj(attrName,0) != null) {
//                            System.out.println("geometry found");
                            IomObject iomGeom = iomObj.getattrobj(attrName,0);
                            if (iomGeom != null) {
                                if (iomGeom.getobjecttag().equals(COORD)) {
                                    attrDesc.setBinding(Point.class);
                                } else if (iomGeom.getobjecttag().equals(MULTICOORD)) {
                                    attrDesc.setBinding(MultiPoint.class);
                                } else if (iomGeom.getobjecttag().equals(POLYLINE)) {
                                    attrDesc.setBinding(LineString.class);
                                } else if (iomGeom.getobjecttag().equals(MULTIPOLYLINE)) {
                                    attrDesc.setBinding(MultiLineString.class);
                                } else if (iomGeom.getobjecttag().equals(MULTISURFACE)) {
                                    int surfaceCount=iomGeom.getattrvaluecount("surface");
                                    if(surfaceCount==1) {
                                        /* Weil das "Schema" anhand des ersten IomObjektes erstellt wird,
                                         * kann es vorkommen, dass Multisurfaces mit mehr als einer Surface nicht zu einem Multipolygon umgewandelt werden,
                                         * sondern zu einem Polygon. Aus diesem Grund wird immer das MultiPolygon-Binding verwendet. */
                                        attrDesc.setBinding(MultiPolygon.class);
                                    } else if (surfaceCount>1) {
                                        attrDesc.setBinding(MultiPolygon.class);
                                    }
                                } else {
                                    // Siehe Kommentar oben. Ist das sinnvoll? Resp. funktioniert das wenn es andere Strukturen gibt? Diese könnte man nach JSON
                                    // umwandeln und als String behandeln.
                                    // Was passiert in der Logik, falls keine Geometrie gesetzt ist?
                                    attrDesc.setBinding(Point.class);
                                }
                                if (defaultSrsId != null) {
                                    attrDesc.setSrId(defaultSrsId);
                                }
                                attrDesc.setGeometry(true);
                            }
                        } else {
                            attrDesc.setBinding(String.class);
                        }
                        attrDesc.setAttributeName(attrName);
                        attrDescs.add(attrDesc);
                    }
                }
            }
            
            if (headerRow == null) {
                workbook = new XSSFWorkbook(); 
                sheet = workbook.createSheet();
                headerRow = sheet.createRow(0);
                
                int cellnum = 0;
                for (ExcelAttributeDescriptor attrDesc : attrDescs) {
                    Cell cell = headerRow.createCell(cellnum++);
                    cell.setCellValue(attrDesc.getAttributeName());
                    
//                    if(obj instanceof String)
//                         cell.setCellValue((String)obj);
//                     else if(obj instanceof Integer)
//                         cell.setCellValue((Integer)obj);

                }
            }
             
            
            // Ich glaube ich muss sicherstellen, dass die Reihenfolge stimmt.
            // Beim Header muss ich attrCellNums Map erstellen. Key ist attrName, value cell nr.
            
            Row row = sheet.createRow(sheet.getLastRowNum()+1);
            int cellnum = 0;
            for (ExcelAttributeDescriptor attrDesc : attrDescs) {
                String attrName = attrDesc.getAttributeName();
                
                // if else trallala 
                // -typen
                // -null
                String attrValue = iomObj.getattrvalue(attrName);
                Cell cell = row.createCell(cellnum++);
                cell.setCellValue(attrValue);


            }
            
        }
    }
    
//    private int getRowNum() {
//        return rowNum++;
//    }

    @Override
    public void close() throws IoxException {
        try {
            FileOutputStream out = new FileOutputStream(outputFile);
            workbook.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
            throw new IoxException(e.getMessage());
        }
    }

    @Override
    public IomObject createIomObject(String arg0, String arg1) throws IoxException {
        return null;
    }

    @Override
    public void flush() throws IoxException {        
    }

    @Override
    public IoxFactoryCollection getFactory() throws IoxException {
        return null;
    }

    @Override
    public void setFactory(IoxFactoryCollection arg0) throws IoxException {        
    }

}
