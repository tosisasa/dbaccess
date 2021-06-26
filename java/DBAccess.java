package tools;

import java.sql.*;
import java.util.ArrayList;

import oracle.jdbc.*;

public class DBAccess {

	Connection conn;

	public Connection connectDB(String schema) throws SQLException,
			ClassNotFoundException {

		if (schema.equals("識別用")) {
			Class.forName("oracle.jdbc.driver.OracleDriver");
			conn = DriverManager.getConnection(
					"jdbc:oracle:thin:@hostname:1521:orcl", "ユーザー名", "パスワード");
		}
		return conn;
	}

	public void closeDB(boolean isCommit) {
		try {
			if (isCommit) {
				conn.commit();
			} else {
				conn.rollback();
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}
		try {
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}

	public ArrayList selectDBOnTrans(String sql, String schema)
			throws SQLException, ClassNotFoundException {
		Statement statement = conn.createStatement();
		ResultSet resultSet = statement.executeQuery(sql);

		ResultSetMetaData metaData = resultSet.getMetaData();
		int columnCount = metaData.getColumnCount();

		while (resultSet.next()) {
			// TODO

			for (int i = 0; i < columnCount; i++) {
				Object obj = resultSet.getObject(i);
			}

		}
		resultSet.close();

		ArrayList list = new ArrayList();
		return list;
	}

	public void updateDB(String sql, String schema) {
	}

	public boolean updateDBOnTrans(String sql, boolean message) {
		boolean normal = true;
		try {
			Statement statement = conn.createStatement();
			statement.executeUpdate(sql);
			statement.close();
		} catch (SQLException e) {
			normal = false;
			e.printStackTrace();
		}
		return normal;

	}

}
