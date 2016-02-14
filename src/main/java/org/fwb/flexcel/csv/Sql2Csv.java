package org.fwb.flexcel.csv;

import java.io.IOException;
import java.io.Writer;
import java.sql.ResultSet;
import java.sql.SQLException;

import au.com.bytecode.opencsv.CSVWriter;

/** @deprecated example-only */
public class Sql2Csv {
	/** @deprecated static utilities only */
	@Deprecated
	private Sql2Csv() { }
	
	public static void rs2csv(
			ResultSet rs, Writer w, boolean header)
			throws SQLException, IOException {
		CSVWriter csv = new CSVWriter(w); try {
			csv.writeAll(rs, header);
		} finally {
			csv.close();
		}
	}
}
