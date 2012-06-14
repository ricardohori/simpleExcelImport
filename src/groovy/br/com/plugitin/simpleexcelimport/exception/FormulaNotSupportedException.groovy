package br.com.plugitin.simpleexcelimport.exception

class FormulaNotSupportedException extends RuntimeException {
	private def tabName
	private def columnLetter
	private def columnLine

    public FormulaNotSupportedException(Throwable cause, tabName, columnLetter, columnLine) {
        super(cause)
        this.tabName = tabName
        this.columnLetter = columnLetter
        this.columnLine = columnLine
    }

	@Override
	public String getMessage() {
		return "The formula could not be resolved within the sheet into the tab '${tabName}' at column '${columnLetter} and line ${columnLine}'"
	}
}
