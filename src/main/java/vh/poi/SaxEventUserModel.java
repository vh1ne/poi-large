package vh.poi;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import java.io.FileInputStream;
import java.io.InputStream;
import java.time.Duration;
import java.time.Instant;
import java.util.Iterator;

public class SaxEventUserModel {

	public void processSheets(String filename) throws Exception {
		Instant start = Instant.now();
		var stream = new FileInputStream(filename);
		IOUtils.setByteArrayMaxOverride(1000000000);

		OPCPackage pkg = OPCPackage.open(stream);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = (SharedStringsTable) r.getSharedStringsTable();
		XMLReader parser = fetchSheetParser(sst);
		Iterator<InputStream> sheets = r.getSheetsData();
		while (sheets.hasNext()) {
			System.out.println("Processing new sheet:\n");
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
			System.out.println("");
		}
		Instant end = Instant.now();
		System.out.println(printExecutionTime(start,end));
	}


	public static String printExecutionTime(Instant start, Instant end)
	{
		return "Program  executed  in "+  (float) Duration.between(start, end).toMillis() / 1000  + " seconds." ;
	}

	public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException, ParserConfigurationException {
		XMLReader parser = XMLHelper.newXMLReader();
		ContentHandler handler = new SheetHandler(sst);
		parser.setContentHandler(handler);
		return parser;
	}

	/**
	 * See org.xml.sax.helpers.DefaultHandler javadocs
	 */
	private static class SheetHandler extends DefaultHandler {
		private SharedStringsTable sst;
		private String lastContents;
		private boolean nextIsString;

		private SheetHandler(SharedStringsTable sst) {
			this.sst = sst;
		}

		@Override
		public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
			// c => cell
			if (name.equals("c")) {
				// Print the cell reference
				System.out.print(attributes.getValue("r") + " - ");
				// Figure out if the value is an index in the SST
				String cellType = attributes.getValue("t");
				if (cellType != null && cellType.equals("s")) {
					nextIsString = true;
				} else {
					nextIsString = false;
				}
			}
			// Clear contents cache
			lastContents = "";
		}

		@Override
		public void endElement(String uri, String localName, String name) throws SAXException {
			// Process the last contents as required.
			// Do now, as characters() may be called more than once
			if (nextIsString) {
				int idx = Integer.parseInt(lastContents);
				lastContents = sst.getItemAt(idx).getString();
				nextIsString = false;
			}
			// v => contents of a cell
			// Output after we've seen the string contents
			if (name.equals("v")) {
				System.out.println(lastContents);
			}
		}

		@Override
		public void characters(char[] ch, int start, int length) {
			lastContents += new String(ch, start, length);
		}
	}

}
