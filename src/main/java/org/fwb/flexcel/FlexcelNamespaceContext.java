package org.fwb.flexcel;

import java.util.*;

import javax.xml.XMLConstants;
import javax.xml.namespace.NamespaceContext;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.collect.ImmutableMap;

class FlexcelNamespaceContext implements NamespaceContext {
	private static final Logger LOG = LoggerFactory.getLogger(FlexcelNamespaceContext.class);
	
	public static final String
		NS_XL = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
		NS_R = "http://schemas.openxmlformats.org/package/2006/relationships",
		NS_XP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
		NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
	
	private static final Map<String, String> PREFIX_TO_URI = ImmutableMap.<String, String>builder()
		.put(XMLConstants.XML_NS_PREFIX, XMLConstants.XML_NS_URI)
		.put(XMLConstants.XMLNS_ATTRIBUTE, XMLConstants.XMLNS_ATTRIBUTE_NS_URI)
		.put("xl", NS_XL)
		.put("r", NS_R)
		.put("xp", NS_XP)
		.put("vt", NS_VT).build();
	
	public static final NamespaceContext NSCTX = new FlexcelNamespaceContext();
	
	private FlexcelNamespaceContext() { }
	
	@Override
	public String getNamespaceURI(String prefix) {
		String retVal = PREFIX_TO_URI.get(prefix);
		if (retVal == null)
			retVal = XMLConstants.NULL_NS_URI;
		return retVal;
	}
	/**
	 * returns the first uri returned by getPrefixes(uri), or null if there is no such prefix
	 */
	@Override
	public String getPrefix(String uri) {
		Iterator<String> i = getPrefixes(uri);
		if (i.hasNext())
			return i.next();
		else
			return null;
	}
	/**
	 * @deprecated TODO always throws UnsupportedOperationException
	 */
	@Override
	public Iterator<String> getPrefixes(String uri) throws UnsupportedOperationException {
		UnsupportedOperationException e = new UnsupportedOperationException("todo");
		LOG.error("called getPrefixes(" + uri + ")", e);
		throw e;
	}
}
