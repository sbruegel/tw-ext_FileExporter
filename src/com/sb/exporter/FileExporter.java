package com.sb.exporter;

import java.io.File;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Iterator;

import com.itextpdf.text.Document;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.pdf.codec.Base64;
import com.thingworx.entities.utils.ThingUtilities;
import com.thingworx.logging.LogUtilities;
import com.thingworx.metadata.annotations.ThingworxServiceDefinition;
import com.thingworx.metadata.annotations.ThingworxServiceParameter;
import com.thingworx.metadata.annotations.ThingworxServiceResult;
import com.thingworx.things.repository.FileRepositoryThing;
import com.thingworx.types.BaseTypes;
import com.thingworx.types.InfoTable;
import com.thingworx.types.primitives.IPrimitiveType;

import ch.qos.logback.classic.Logger;

public class FileExporter {
	
	protected static final Logger _Logger = LogUtilities.getInstance().getApplicationLogger(FileExporter.class);

	public FileExporter() {
		// TODO Auto-generated constructor stub
	}
	
	@SuppressWarnings("rawtypes")
	@ThingworxServiceDefinition(
            name = "ExportInfotableAsPdf",
            description = ""
    )
    @ThingworxServiceResult(
            name = "result",
            description = "",
            baseType = "STRING"
    )
    public String ExportInfotableAsPdf(	@ThingworxServiceParameter(name = "infotable", description = "Infotable you want to export", baseType = "INFOTABLE", aspects = { "isRequired:true"}) InfoTable infotable,
    									@ThingworxServiceParameter(name = "fileName", description = "Optinal name of the file, if this parameter is empty file name is 'Export-yyyy-MM-dd_HH-mm-ss'", baseType = "STRING", aspects = { "isRequired:false"}) String fileName,
            							@ThingworxServiceParameter(name = "fileRepository", description = "Reference to a File Repository Thing where the exported file will be saved", baseType = "THINGNAME", aspects = { "isRequired:true", "thingTemplate:FileRepository" }) String fileRepository) 
            							throws Exception {
		
	    
		int columnSize = infotable.getRow(0).size();
        Document document = new Document();
        
        String exportFileName = fileName;
        
        if(exportFileName == "" || exportFileName == null) {
        	DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
    	    Calendar cal = Calendar.getInstance();
        	exportFileName = "Export-" + dateFormat.format(cal.getTime());
        }
        
        // Search the file Repo given by Service
        FileRepositoryThing fileRepositoryThingRef = (FileRepositoryThing)ThingUtilities.findThing(fileRepository);
        
        File file = new File(fileRepositoryThingRef.getRootPath() + File.separator + exportFileName + ".pdf");

        PdfWriter.getInstance(document, new FileOutputStream(file));
        document.open();
        PdfPTable table = new PdfPTable(columnSize);
        table.setSplitLate(false);
        Iterator<String> FieldNames = infotable.getRow(0).keySet().iterator();

        // Interate through Datashape
        while(FieldNames.hasNext()) {
            String name = (String)FieldNames.next();
            table.addCell(name);
        }

        for(int i = 0; i < infotable.getRowCount(); ++i) {
            Iterator<IPrimitiveType> content = infotable.getRow(i).values().iterator();

            while(content.hasNext()) {
                IPrimitiveType value = (IPrimitiveType)content.next();
                if(value.getBaseType() == BaseTypes.IMAGE) {
                	Image img = Image.getInstance(Base64.decode(value.getStringValue()));
                	table.addCell(img);
                }
                else
                	table.addCell(value.getStringValue());
            }
        }

        document.add(table);
        document.close();
        return fileRepositoryThingRef.GetFileListingWithLinks("/", exportFileName + ".pdf").getRow(0).getStringValue("downloadLink");
        
		//return exp.ExportInfotableAsPdf(infotable, fileRepository);
	}

    /*@ThingworxServiceDefinition(
            name = "ExportInfotableAsExcel",
            description = ""
    )
    @ThingworxServiceResult(
            name = "result",
            description = "",
            baseType = "STRING"
    )
    public String ExportInfotableAsExcel(	@ThingworxServiceParameter(name = "infotable",description = "",baseType = "INFOTABLE") InfoTable infotable,
            								@ThingworxServiceParameter(name = "fileRepository",description = "Reference to a File Repository Thing where the exported file will be saved",baseType = "THINGNAME", aspects = { "isRequired:true", "thingTemplate:FileRepository" }) String fileRepository) throws Exception {
		ExporterUtils exp = new ExporterUtils();
		return exp.ExportInfotableAsExcel(infotable, fileRepository);
	}

    @ThingworxServiceDefinition(
            name = "ExportInfotableAsWord",
            description = ""
    )
    @ThingworxServiceResult(
            name = "result",
            description = "",
            baseType = "STRING"
    )
    public String ExportInfotableAsWord(@ThingworxServiceParameter(name = "infotable",description = "",baseType = "INFOTABLE") InfoTable infotable,
            							@ThingworxServiceParameter(name = "fileRepository",description = "Reference to a File Repository Thing where the exported file will be saved",baseType = "THINGNAME", aspects = { "isRequired:true", "thingTemplate:FileRepository" }) String fileRepository) throws Exception {
		ExporterUtils exp = new ExporterUtils();
		return exp.ExportInfotableAsWord(infotable, fileRepository);
	}*/

}
