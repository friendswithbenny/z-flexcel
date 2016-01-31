package org.fwb.flexcel;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.base.Charsets;
import com.google.common.io.Files;

/**
 * testing utility to create-and-run Flexcel files
 * on a system which has Microsoft Word (TM) installed.
 */
public class FlexcelRuntime {
	private static final Logger logger = LoggerFactory.getLogger(FlexcelRuntime.class);
	
	/** executable command to launch MS Excel (TM) */
	private static final String EXCEL = System.getProperty("org.fwb.flexcel.FlexcelRuntime.EXCEL",
			"excel");
//			"excel.exe");	// explicit for MS Windows (TM)
	private static final int SLEEP_SECONDS = Integer.getInteger("org.fwb.flexcel.FlexcelRuntime.SLEEP_SECONDS", 5);
	
	private static final Runtime RUNTIME = Runtime.getRuntime();
	
	/**
	 * given a query script, executes each query.
	 * saves the results to a temporary XLSX file and opens it.
	 * each resultset is named by its FROM tablename (which apparently turns out to be blank?!) prefixed with rs#.
	 * n.b. this method *requires* the only semicolons in the document be separating queries (i.e. no semicolons in strings nor comments)
	 */
	public static final void queries2excel(File sqls, Connection c) throws FlexcelException,
			IOException, SQLException, InterruptedException {
		String SQLS = Files.toString(sqls, Charsets.UTF_8);
		queries2excel(c, SQLS.split("[;]"));
	}
	public static final void queries2excel(Connection c, String[] sqls) throws FlexcelException,
			SQLException, IOException, InterruptedException {
		Statement s = c.createStatement();
			File f = File.createTempFile("odf", ".xlsx");
			Flexcel x = Flexcel.create(f); try {
				for (int i = 0; i < sqls.length; ++i) {
					if (sqls[i].trim().length() > 0) {
						ResultSet rs;
						try {
							rs = s.executeQuery(sqls[i]);
						} catch (Throwable t) {
							throw new RuntimeException("query failed:\r\n\t" + sqls[i], t);
						}
							ResultSetMetaData rsmd = rs.getMetaData();
							String name = rsmd.getTableName(1);	// TODO figure out the "from" clause's name
							x.dataDump("rs" + i + name, rs, true);
						rs.close();
					}
				}
				x.save();
			} finally {
				x.close();
			}
		s.close();
		
		String[] args = {EXCEL, "/r", f.getCanonicalPath()};
		RUNTIME.exec(args);
		Thread.sleep(SLEEP_SECONDS * 1000);
		System.out.println("delete: " + f.delete());
	}
	
	public static final void viewTempXL(ResultSet rs) throws FlexcelException,
			IOException, InterruptedException {
		File f = File.createTempFile("odf", ".xlsx"); try {
			Flexcel.simpleReport(rs, f);
			String[] args = {EXCEL, "/r", f.getCanonicalPath()};
			RUNTIME.exec(args);
			Thread.sleep(SLEEP_SECONDS * 1000);
		} finally {
			logger.warn("delete: " + f.delete());
		}
	}
}
