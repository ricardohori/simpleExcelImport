package br.com.plugitin.simpleexcelimport.exception

class ColumnNotFoundException extends Exception {
	
	private def tabName
	private def columnLetter
	private def columnName
	
	@Override
	public String getMessage() {
		return "Column header '${columnName}' was not found within the sheet into the tab '${tabName}' at column '${columnLetter}'"
	}
}
