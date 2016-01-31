package org.fwb.flexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;
import javax.xml.transform.Templates;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.sax.SAXTransformerFactory;
import javax.xml.transform.sax.TransformerHandler;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLFilter;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLFilterImpl;

/**
 * this is a self-contained XMLFilter chain
 * whose first filter adds a Date Style,
 * and the second filter simply harvests its ID.
 */
class AddAndGetDateStyle extends XMLFilterImpl {
	private static final Logger LOG = LoggerFactory.getLogger(AddAndGetDateStyle.class);
	
	private static final SAXTransformerFactory STF = (SAXTransformerFactory) SAXTransformerFactory.newInstance();
	
	/**
	 * uses self-contained parser and serializer
	 */
	static final int addAndGetDateStyle(File xml) throws ParserConfigurationException, SAXException, TransformerConfigurationException, IOException {
		AddAndGetDateStyle ads = new AddAndGetDateStyle();
		filterInPlace(xml, ads);
		return ads.getStyleIndex();
	}
	
	private static final void filterInPlace(File srcDst, XMLFilter f)
			throws ParserConfigurationException, SAXException, IOException, TransformerConfigurationException {
		File src = new File(srcDst.getParentFile(), srcDst.getName() + "." + new Date().getTime());
		
		XMLReader parser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
		f.setParent(parser);
		
		if (srcDst.renameTo(src)) {
			InputStream is = new FileInputStream(src);
			OutputStream os = new FileOutputStream(srcDst);
			
			TransformerHandler th = STF.newTransformerHandler();
			th.setResult(new StreamResult(os));
			f.setContentHandler(th);
			f.parse(new InputSource(is));
			is.close();
			os.close();
			
			if (! src.delete())
				LOG.error("filterInPlace(_) unable to delete temporary pre-filter file: {}", src);
		} else
			throw new IOException("unable to rename [" + srcDst + "] to [" + src + "]");
	}
	
	private static final Templates ADD_DATE_STYLE;
	static {
		try {
			InputStream is = AddAndGetDateStyle.class.getResourceAsStream("add-date-style.xsl");
			ADD_DATE_STYLE = ((SAXTransformerFactory) SAXTransformerFactory.newInstance()).newTemplates(new StreamSource(is));	// TransformerConfigurationException
			is.close();															// IOException
		} catch (TransformerConfigurationException e) {
			throw new RuntimeException(e);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}
	
	/* INSANCE */
	private final XMLFilter PAR;
	private int styleIndex = -1;
	private AddAndGetDateStyle() throws TransformerConfigurationException {
		super();
		PAR = ((SAXTransformerFactory) SAXTransformerFactory.newInstance()).newXMLFilter(ADD_DATE_STYLE);
		super.setParent(PAR);
	}
	
	private int getStyleIndex() {
		return styleIndex;
	}
	
	@Override
	public void setParent(XMLReader parent) {
		PAR.setParent(parent);
	}
	@Override
	public void startDocument() throws SAXException {
		super.startDocument();	// SAXException
		styleIndex = -1;
	}
	@Override
	public void startElement(String uri, String localName, String qName, Attributes at) throws SAXException {
		super.startElement(uri, localName, qName, at);	// SAXException
		if ("cellXfs".equals(localName))
			styleIndex = Integer.parseInt(at.getValue("count")) - 1;
	}
}
