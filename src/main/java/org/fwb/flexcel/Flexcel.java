package org.fwb.flexcel;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.io.*;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Templates;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.dom.DOMResult;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.sax.SAXTransformerFactory;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.fwb.dir.ZipDirectory;
import org.fwb.dir.ZipUtility;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import com.google.common.base.Objects;

public class Flexcel implements Closeable {
	private static final Logger LOG = LoggerFactory.getLogger(Flexcel.class);
	
	static final SAXTransformerFactory STF = (SAXTransformerFactory) SAXTransformerFactory.newInstance();
	private static final DocumentBuilderFactory DBF = DocumentBuilderFactory.newInstance();
	private static final XPath XP = XPathFactory.newInstance().newXPath();
	
	private static final String FLEXCEL_DATA_DIR_NAME = "flexcelData";
	
	private static final XPathExpression
		XPATH_RELS_PROPS,
			XPATH_PROPS_NAMEDRANGES,
			XPATH_PROPS_TITLES,
		XPATH_RELS_BOOK,
			XPATH_BOOKRELS_STYLES,
			XPATH_DATE_1904,
			XPATH_MAX_SHEET_ID;
	private static final Templates
		ADD_NAMED_RANGES,
		ADD_DEFINED_NAMES;
	static {
		/* initialize XPath */
		XP.setNamespaceContext(FlexcelNamespaceContext.NSCTX);
		
		/* compile XPathExpressions */
		try {
			XPATH_RELS_PROPS = XP.compile(
					"r:Relationships/r:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties']/@Target");
			XPATH_PROPS_NAMEDRANGES = XP.compile(
					"xp:Properties/xp:HeadingPairs/vt:vector/vt:variant[vt:lpstr = 'Named Ranges']/following-sibling::vt:variant[1]/vt:i4");
			XPATH_PROPS_TITLES = XP.compile(
					"xp:Properties/xp:TitlesOfParts/vt:vector");
			XPATH_RELS_BOOK = XP.compile(
					"r:Relationships/r:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument']/@Target");
			XPATH_BOOKRELS_STYLES = XP.compile(
					"r:Relationships/r:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles']/@Target");
			XPATH_DATE_1904 = XP.compile(
					"xl:workbook/xl:workbookPr/@date1904 = 1");
			XPATH_MAX_SHEET_ID = XP.compile(
					"xl:workbook/xl:sheets/xl:sheet/@sheetId[not(number() < /xl:workbook/xl:sheets/xl:sheet/@sheetId)]");
		} catch (XPathExpressionException e) {
			throw new RuntimeException(e);
		}
		
		/* loading Template resources */
		try {
			InputStream is;
			is = Flexcel.class.getResourceAsStream("add-named-ranges.xsl");
			ADD_NAMED_RANGES = STF.newTemplates(new StreamSource(is));
			is.close();
			
			is = Flexcel.class.getResourceAsStream("add-defined-names.xsl");
			ADD_DEFINED_NAMES = STF.newTemplates(new StreamSource(is));
			is.close();
		} catch (TransformerConfigurationException e) {
			throw new RuntimeException(e);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}
	
	/**
	 * @param xlsx report output File
	 * @param names the names of tables (or result-set tructures, like views) from which to select *
	 */
	public static void dumpTables(File xlsx, Connection c, String... names) throws FlexcelException {
//			throws TransformerException, SQLException, SAXException, IOException, XPathExpressionException, ParserConfigurationException, InterruptedException {
		String[][] namesAndSortColumns = new String[names.length][2];
		for (int i = 0; i < names.length; ++i)
			namesAndSortColumns[i][0] = names[i];
		dumpTablesSorted(xlsx, c, namesAndSortColumns);
		
//		Flexcel x = new Flexcel(emptyWorkbook(), xlsx); try {
//			Statement s = c.createStatement();
//			ResultSet rs;
//			for (String t : names) {
//				rs = s.executeQuery("select * from " + t);
//				x.dataDump(t, rs, true);
//				rs.close();
//			}
//			s.close();
//			
//			x.save();
//		} finally {
//			x.close();
//		}
	}
	/**
	 * 
	 * @param xlsx report output file
	 * @param c database connection
	 * @param namesAndSortColumns
	 */
	public static void dumpTablesSorted(File xlsx, Connection c, String[][] namesAndSortColumns) throws FlexcelException {
//			throws TransformerException, SQLException, SAXException, IOException, XPathExpressionException, ParserConfigurationException, InterruptedException {
		Flexcel x = new Flexcel(emptyWorkbook(), xlsx); try {
			Statement s = c.createStatement(); try {
				ResultSet rs;
				for (String[] kv : namesAndSortColumns) {
					rs = s.executeQuery("select * from " + kv[0] + 
							(kv[1] == null ? "" : " order by " + kv[1]));
					try {
						x.dataDump(kv[0], rs, true);
					} finally {
						rs.close();
					}
				}
			} finally {
				s.close();
			}
			x.save();
		} catch (SQLException e) {
			throw new FlexcelException(e);
		} finally {
			closeWrapped(x);
		}
	}
	
	/**
	 * outputs the results of a single query to an xlsx File.
	 * does NOT close the ResultSet when done.
	 * 
	 * @param rs data
	 * @param xlsx target File
	 */
	public static void simpleReport(ResultSet rs, File xlsx) throws FlexcelException {
//			throws TransformerException, SQLException, SAXException, IOException, XPathExpressionException, ParserConfigurationException {
		Flexcel x = new Flexcel(emptyWorkbook(), xlsx); try {
			x.dataDump("test_data", rs, true);
			x.save();
		} finally {
			closeWrapped(x);
		}
	}
	
	public static Flexcel create(File xlsx) throws FlexcelException {
		if (xlsx.exists())
			throw new IllegalArgumentException("can't create new workbook at existing location: " + xlsx);
		
		return new Flexcel(emptyWorkbook(), xlsx);
	}
	public static Flexcel create(File xlsx, File template) throws FlexcelException {
		if (xlsx.exists())
			throw new IllegalArgumentException("can't create new workbook at existing location: " + xlsx);
		
		return createOrOpen(xlsx, template);
	}
	public static Flexcel open(File xlsx) throws FlexcelException {
		return createOrOpen(xlsx, xlsx);
	}
	private static Flexcel createOrOpen(File xlsx, File template) throws FlexcelException {
		try {
			InputStream is = new FileInputStream(template); try {
				return new Flexcel(is, xlsx);
			} finally {
				is.close();
			}
		} catch (IOException e) {
			throw new FlexcelException(e);
		}
	}
	
	private static InputStream emptyWorkbook() {
		return Flexcel.class.getResourceAsStream("empty.xlsx");
	}
	
	/** close the given closeable, wrapping IOException with FlexcelException */
	private static void closeWrapped(Closeable c) throws FlexcelException {
		try {
			c.close();
		} catch (IOException e) {
			throw new FlexcelException("failed to close: " + c, e);
		}
	}
	
	/**
	 * helper function to convert a column number to an Excel column name
	 * @param i	0-based index
	 * @return	the excel column name (alpha)
	 */
	private static String columnIndex(int i) {
		if (i < 26)
			return Character.toString((char) ('A' + i));
		else
			return columnIndex(i/26 - 1) + columnIndex(i%26);
	}
	
	/* INSTANCE */
	/**
	 * File structure of necessary files (ignoring many)
	 */
	public final File
		XLSX;
	public final ZipDirectory ROOT;
	public final File
		PROPS_APP,
		XL_ROOT,
			XL_BOOK,
				XL_BOOK_RELS,
			XL_STYLES,
			FLEXCEL_DATA,
			XL_RELS,
		ROOT_RELS,
			ROOT_DOT_RELS;
	
	/**
	 * in-memory load of certain files which must be sequentially manipulated
	 * these Objects must (finally) be persisted/serialized for any Flexcel manipulation to be valid
	 */
	public final Document DOM_BOOK, DOM_RELS, DOM_PROPS;
	
	/**
	 * specific objects in each DOM which will have new members appended
	 */
	public final Element SHEETS, NAMES, RELATIONSHIPS, TITLES, NAMEDRANGES;
	
	/* properties of the Flexcel */
	public final boolean DATE_1904;
	public final int DATE_STYLE_INDEX;
	private int maxSheetId;
	
	/**
	 * uses an InputStream as xlsx template content.
	 * the InputStream is NOT closed upon completion
	 */
	private Flexcel(InputStream template, File xlsx) throws FlexcelException {
//			throws SAXException, IOException, XPathExpressionException, ParserConfigurationException, TransformerException {
		LOG.debug("start <init>({}, {})", template, xlsx);
		XLSX = xlsx;
		
		try {
			DocumentBuilder DB = DBF.newDocumentBuilder();
			String p;
			DOMResult r;
			ROOT = new ZipDirectory(xlsx, "Flexcel", ".flxl.zip.tmpdir", xlsx.getParentFile());
				ROOT_RELS = new File(ROOT, "_rels");
					ROOT_DOT_RELS = new File(ROOT_RELS, ".rels");
					ZipUtility.unzip(template, ROOT);
					Document rels = DB.parse(ROOT_DOT_RELS);	// not manipulated, so not persisted, so not stored as instance state
					p = (String) XPATH_RELS_PROPS.evaluate(rels, XPathConstants.STRING);
					
					PROPS_APP = new File(ROOT, p);
						r = new DOMResult();
						ADD_NAMED_RANGES.newTransformer().transform(new StreamSource(PROPS_APP), r);
						DOM_PROPS = (Document) r.getNode();
							TITLES = (Element) XPATH_PROPS_TITLES.evaluate(DOM_PROPS, XPathConstants.NODE);
							NAMEDRANGES = (Element) XPATH_PROPS_NAMEDRANGES.evaluate(DOM_PROPS, XPathConstants.NODE);
					p = (String) XPATH_RELS_BOOK.evaluate(rels, XPathConstants.STRING);
					XL_BOOK = new File(ROOT, p);
						r = new DOMResult();
						ADD_DEFINED_NAMES.newTransformer().transform(new StreamSource(XL_BOOK), r);
						DOM_BOOK = (Document) r.getNode();
							SHEETS = (Element) DOM_BOOK.getElementsByTagName("sheets").item(0);
							NAMES = (Element) DOM_BOOK.getElementsByTagName("definedNames").item(0);
				XL_ROOT = XL_BOOK.getParentFile();
					XL_RELS = new File(XL_ROOT, "_rels");
						XL_BOOK_RELS = new File(XL_RELS, XL_BOOK.getName() + ".rels");
							DOM_RELS = DB.parse(XL_BOOK_RELS);
								RELATIONSHIPS = DOM_RELS.getDocumentElement();
					p = (String) XPATH_BOOKRELS_STYLES.evaluate(DOM_RELS, XPathConstants.STRING);	// XPathExpressionException
					XL_STYLES = new File(XL_ROOT, p);
					FLEXCEL_DATA = new File(XL_ROOT, FLEXCEL_DATA_DIR_NAME);
			
			// flag for Excel recalculation
			((Element) DOM_BOOK.getElementsByTagName("calcPr").item(0)).setAttribute("fullCalcOnLoad", "true");
			
			FLEXCEL_DATA.mkdir();
			
			DATE_1904 = (Boolean) XPATH_DATE_1904.evaluate(DOM_BOOK, XPathConstants.BOOLEAN);
			DATE_STYLE_INDEX = AddAndGetDateStyle.addAndGetDateStyle(XL_STYLES);					// ParserConfigurationException, SAXException, TransformerConfigurationException, IOException
			maxSheetId = ((Double) XPATH_MAX_SHEET_ID.evaluate(DOM_BOOK, XPathConstants.NUMBER)).intValue();
			LOG.debug("end <init>(_)");
		} catch (Exception e) {
			throw new FlexcelException(e);
		}
	}
	
	private String getRelativePath(File f) {
		return f.toURI().relativize(XL_ROOT.toURI()).getPath();
//		String retVal = f.getCanonicalPath();
//		retVal = retVal.substring(XL_ROOT.getCanonicalPath().length());
//		if (retVal.startsWith(File.separator))
//			retVal = retVal.substring(File.separator.length());
//		return retVal;
	}
	
	/**
	 * adds a worksheet of data to the workbook
	 * the worksheet is added in the "odf" directory
	 * 
	 * the worksheet is added to the workbook "table of contents" (workbook.xml) and its relationships (workbook.xml.rels) DOM objects
	 * a named range is likewise added to the workbook DOM
	 */
	/*
	 * TODO validate input name: max-length? legal chars?
	 * 
	 * TODO	bold the header rows of data dumps
		ActiveWindow.SplitRow = 1
		ActiveWindow.FreezePanes = True
		Rows(1).Font.Bold = True
	 */
	public void dataDump(String name, ResultSet rs, boolean visible) throws FlexcelException {
//			throws TransformerConfigurationException, SQLException, SAXException, IOException {
		LOG.debug("start dataDump({}, {}, {})", name, rs, visible);
		
//		name = "odf_" + name;
		int columnCount;
		try {
			columnCount = rs.getMetaData().getColumnCount();
		} catch (SQLException e) {
			throw new FlexcelException(e);
		}
		
		/* data */
		File dst = new File(FLEXCEL_DATA, name + ".xml");
		int records = Sql2WorksheetXml.sql2xl(rs, DATE_STYLE_INDEX, dst);
		
		synchronized (this) {
			/* sheet */
			// rels
			Element rel = DOM_RELS.createElement("Relationship");
			rel.setAttribute("Id", name);
			rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
			rel.setAttribute("Target", getRelativePath(dst));
			RELATIONSHIPS.appendChild(rel);
			
			// book
			Element sheet = DOM_BOOK.createElement("sheet");
			sheet.setAttribute("name", name);
			sheet.setAttribute("sheetId", "" + (++ maxSheetId));
			sheet.setAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships",
					"r:id", name);
			if (! visible)
				sheet.setAttribute("state", visible ? "visible" : "veryHidden");
//			sheet.setAttribute("state", visible ? "visible" : "veryHidden");
			SHEETS.appendChild(sheet);
			
			/* name */
			// book
			Element dName = DOM_BOOK.createElement("definedName");
			dName.setAttribute("name", name);
			dName.appendChild(DOM_BOOK.createTextNode(
					name + "!$A$1:$" + columnIndex(columnCount - 1) + "$" + (records+1)));
			NAMES.appendChild(dName);
			
			// titles
			Node newNode = DOM_PROPS.createElementNS(FlexcelNamespaceContext.NS_VT, "vt:lpstr");
			newNode.appendChild(DOM_PROPS.createTextNode(name));
			TITLES.appendChild(newNode);
			TITLES.setAttribute("size", "" + (1 + Integer.parseInt(TITLES.getAttribute("size"))));
			
			// heading-pairs
			Node val = NAMEDRANGES.getFirstChild();
			NAMEDRANGES.replaceChild(DOM_PROPS.createTextNode("" + (1 + Integer.parseInt(val.getTextContent()))), val);
		}
		LOG.debug("end dataDump(_)");
	}
	
	public void save() throws FlexcelException {
//			throws TransformerException, IOException {
		LOG.debug("start save()");
		try {
			serializeDom(DOM_RELS, XL_BOOK_RELS);
			serializeDom(DOM_BOOK, XL_BOOK);
			serializeDom(DOM_PROPS, PROPS_APP);
			
			ROOT.zip();
		} catch (Exception e) {
			throw new FlexcelException(e);
		}
		LOG.debug("end save()");
	}
	
	private static void serializeDom(Document d, File f) throws TransformerException, IOException {
		Transformer t = STF.newTransformer();	// identity, for serialization
		OutputStream os = new FileOutputStream(f); try {
			t.transform(new DOMSource(d), new StreamResult(f));
		} finally {
			os.close();
		}
	}
	
//	/** @deprecated close() should always be in the finally block, save() should ~never be there */
//	public void saveAndClose() throws IOException {
//		save();
//		close();
//	}
	
	/** delete the temporary directory */
	@Override
	public void close() throws IOException {
		ROOT.close();
	}
	
	void debug() {
		LOG.debug(Objects.toStringHelper(Flexcel.class)
				.add("XLSX", XLSX)
				.add("ROOT", ROOT)
				.add("XL_ROOT", XL_ROOT)
				.add("XL_BOOK", XL_BOOK)
				.add("XL_BOOK_RELS", XL_BOOK_RELS)
				.add("XL_STYLES", XL_STYLES)
				.add("FLEXCEL_DATA", FLEXCEL_DATA)
				
				.add("DOM_BOOK", DOM_BOOK)
				.add("DOM_RELS", DOM_RELS)
				
				.add("SHEETS", SHEETS)
				
				.add("NAMES", NAMES)
				
				.add("RELATIONSHIPS", RELATIONSHIPS)
				.add("DATE_STYLE_INDEX", DATE_STYLE_INDEX)
				.add("maxSheetId", maxSheetId)
				.toString());
	}
	
	@Override
	public String toString() {
		return Objects.toStringHelper(Flexcel.class)
				.add("XLSX", XLSX)
				.add("ROOT", ROOT)
				.toString();
	}
}
