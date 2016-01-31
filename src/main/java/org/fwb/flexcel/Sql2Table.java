package org.fwb.flexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.Collection;

import javax.swing.text.html.HTML.Tag;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.sax.SAXTransformerFactory;
import javax.xml.transform.sax.TransformerHandler;
import javax.xml.transform.stream.StreamResult;

import org.fwb.sql.RecordList;
import org.fwb.sql.RecordList.StringRecordList;
import org.xml.sax.ContentHandler;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.AttributesImpl;
import org.xml.sax.SAXException;

import com.google.common.base.Function;
import com.google.common.collect.Collections2;

/**
 * utility to convert relational sql results (jdbc ResultSet) into html-style, strict XML data dump.
 * 
 * @deprecated this class is no longer needed, as it provides no real speed gains over using SQL2XML with a to-table transform
 * @see org.fwb/???/SQL2XML.toTable(_)
 */
class Sql2Table {
	/**
	 * unbelievable that these constants don't exist, and protected constructors require strict subclasses :(
	 */
	public static final Tag
		THEAD = new Tag("thead") { },
		TBODY = new Tag("tbody") { };
	
	public static final void toTable(ResultSet rs, File f, boolean header, Function<Object, String> serializer)
			throws SQLException, SAXException, IOException {
		OutputStream os = new FileOutputStream(f);
		toTable(rs, os, header, serializer);
		os.close();
	}
	public static final void toTable(ResultSet rs, OutputStream os, boolean header, Function<Object, String> serializer)
			throws SQLException, SAXException {
		TransformerHandler th;
		try {
			th = ((SAXTransformerFactory) SAXTransformerFactory.newInstance()).newTransformerHandler();
		} catch (TransformerConfigurationException e) {
			throw new RuntimeException("never happens", e);
//		} catch (TransformerFactoryConfigurationError e) {
//			throw new RuntimeException("never happens", e);
		}
		th.setResult(new StreamResult(os));
		toTable(rs, th, header, serializer);
	}
	/**
	 * @return	the number of data records written (not including the header)
	 * @throws SAXException thrown by {@code sax}
	 * @throws SQLException thrown by {@code rs}
	 */
	public static final int toTable(ResultSet rs, ContentHandler sax, boolean header, Function<Object, String> serializer)
			throws SAXException, SQLException {
		Collection<String> rl = serializer == null ? new StringRecordList(rs) : Collections2.transform(new RecordList<Object>(rs), serializer);
		
		start(sax, Tag.TABLE);
			if (header) {
				start(sax, THEAD);
					header(sax, rs.getMetaData());
				end(sax, THEAD);
			}
			
			start(sax, TBODY);
				int retVal;
				for (retVal = 0; rs.next(); ++retVal)
					record(sax, rl);
			end(sax, TBODY);
		end(sax, Tag.TABLE);
		
		return retVal;
	}
	
	/**
	 * writes out a header row, including SQL types as "type" attribute
	 */
	private static void header(ContentHandler sax, ResultSetMetaData rsmd) throws SQLException, SAXException {
		start(sax, Tag.TR);
			char[] s;
			AttributesImpl a;
			for (int i = 0; i < rsmd.getColumnCount(); ++i) {
				a = new AttributesImpl();
				s = rsmd.getColumnName(i+1).toCharArray();
				a.addAttribute("", "type", "type", "CDATA", rsmd.getColumnTypeName(i+1));
				sax.startElement("", Tag.TH.toString(), Tag.TH.toString(), a);
				sax.characters(s, 0, s.length);
				end(sax, Tag.TD);
			}
		end(sax, Tag.TR);
	}
	/**
	 * writes out a record
	 */
	private static void record(ContentHandler sax, Collection<String> ls) throws SAXException {
		start(sax, Tag.TR);
			for (String s : ls) {
				start(sax, Tag.TD);
				if (s != null)
					sax.characters(s.toCharArray(), 0, s.length());
				end(sax, Tag.TD);
			}
		end(sax, Tag.TR);
	}
	
	/* shortcuts for empty namespace and attributes */
	private static final Attributes ATT_EMPTY = new AttributesImpl();
	private static void start(ContentHandler sax, Tag t) throws SAXException {
		sax.startElement("", t.toString(), t.toString(), ATT_EMPTY);
	}
	private static void end(ContentHandler sax, Tag t) throws SAXException {
		sax.endElement("", t.toString(), t.toString());
	}
}
