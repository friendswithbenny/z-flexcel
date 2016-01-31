package org.fwb.flexcel;

import java.sql.Clob;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Date;
import java.io.*;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;
import javax.xml.transform.Result;
import javax.xml.transform.Templates;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.sax.TransformerHandler;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import com.google.common.base.Function;

class Sql2WorksheetXml {
	private static final Logger LOG = LoggerFactory.getLogger(Sql2WorksheetXml.class);
	
	private static final Function<Object, String>
		SERIALIZER = new Function<Object, String>() {
			@Override
			public String apply(Object o) {
				if (o == null)
					return "";
				if (o instanceof Date)
					return Long.toString(((Date) o).getTime());
				if (o instanceof Clob)
					try {
						return ((Clob) o).getSubString(1, (int) Math.min(Integer.MAX_VALUE, ((Clob) o).length()));
					} catch (SQLException e) {
						throw new RuntimeException(e);
					}
				
				return o.toString();
			}
		};
	
	private static final Templates TABLE2XL;
	static {
		try {
			InputStream is = Sql2WorksheetXml.class.getResourceAsStream("table2excel.xsl");
			TABLE2XL = Flexcel.STF.newTemplates(new StreamSource(is));
			is.close();
		} catch (TransformerConfigurationException e) {
			throw new RuntimeException(e);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}
	
	/**
	 * dumps a single tabular relational result into an ooxml worksheet xml file
	 * 
	 * @param rs single tabular relational result (or all but first are ignored)
	 * @param dateStyleIndex ooxml-specific implementation detail (required, unique per sheet)
	 * @param dst destination worksheet xml file
	 * @return the number of data records written (not including the header)
	 */
	static int sql2xl(ResultSet rs, int dateStyleIndex, File dst) throws FlexcelException {
//			throws TransformerConfigurationException, IOException, SAXException, SQLException {
		LOG.debug("start sql2xl({}, {}, {} ({}))", rs, dateStyleIndex, dst, dst.exists() ? "exists" : "DNE");
		
		TransformerHandler th;
		try {
			th= Flexcel.STF.newTransformerHandler(TABLE2XL);
		} catch (TransformerConfigurationException e) {
			throw new RuntimeException("never happens", e);
		}
		LOG.trace("got th: {}", th);
		th.getTransformer().setParameter("dateStyleIndex", Integer.toString(dateStyleIndex));
		LOG.trace("constructing new StreamResult({}) (parent {})", dst, dst.getParentFile().exists() ? "exists" : "DNE");
		
		try {
			/*
			 * OutputStream used explicitly instead of directly using File
			 * to curb known bugs in many XML implementations.
			 */
//			Result r = new StreamResult(dst);
//			LOG.trace("constructed StreamResult(File): {}", r);
			OutputStream os = new FileOutputStream(dst); try {
				Result r = new StreamResult(os);
				LOG.trace("reconstructed StreamResult(OutputStream): {}", r);
				
				th.setResult(r);
				LOG.trace("successfully set result");
				
				th.startDocument();
				@SuppressWarnings("deprecation")
				int retVal = Sql2Table.toTable(rs, th, true, SERIALIZER);
				LOG.debug("end sql2xl(_): {}", retVal);
				th.endDocument();
				
				return retVal;
			} finally {
				os.close();
			}
		} catch (IOException e) {
			throw new FlexcelException(e);
		} catch (SAXException e) {
			throw new FlexcelException(e);
		} catch (SQLException e) {
			throw new FlexcelException(e);
		}
	}
	
	/**
	 * helper method that uses pre-cached data-dump xml file, rather than a live query
	 * 
	 * @deprecated no longer used?
	 */
	@SuppressWarnings("unused")
	private static void sql2xl(File src, int dateStyleIndex, File dst)
			throws TransformerConfigurationException, ParserConfigurationException, SAXException, IOException {
		TransformerHandler xs = Flexcel.STF.newTransformerHandler(TABLE2XL);
		xs.getTransformer().setParameter("dateStyleIndex", "" + dateStyleIndex);
		Result sr = new StreamResult(dst);
		xs.setResult(sr);
		
		XMLReader r = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
		r.setContentHandler(xs);
		
		InputStream is = new FileInputStream(src);
		r.parse(new InputSource(is));
		is.close();
	}
}
