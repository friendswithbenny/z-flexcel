package org.fwb.flexcel;

/**
 * global root exception class used to wrap any underlying exceptions thrown in Flexcel implementation:
 * IOException, TransformerException/SAXException, SQLException, etc.
 */
public class FlexcelException extends Exception {
	/** default */
	private static final long serialVersionUID = 1L;

	public FlexcelException(Throwable cause) {
		super(cause);
	}
	public FlexcelException(String msg) {
		super(msg);
	}
	public FlexcelException(String msg, Throwable cause) {
		super(msg, cause);
	}
}
