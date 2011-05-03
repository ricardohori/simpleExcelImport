package br.com.plugit.simpleexcelimport.exception

class NotADateColumnException extends Exception {
	@Override
	public String getMessage() {
		return "Column incorrectly specified as Date!"
	}
}
