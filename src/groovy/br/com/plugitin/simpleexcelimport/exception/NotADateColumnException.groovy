package br.com.plugitin.simpleexcelimport.exception

class NotADateColumnException extends RuntimeException {
	def tabName
	def columnLetter
	def rowNumber
	
	@Override
	public String getMessage() {
		return "Invalid date at cell ${rowNumber}, column ${columnLetter} within tab ${tabName}"
	}
}
