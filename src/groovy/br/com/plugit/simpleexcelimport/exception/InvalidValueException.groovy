package br.com.plugit.simpleexcelimport.exception

class InvalidValueException extends Exception {
	@Override
	public String getMessage() {
		return "Invalid value(s)!"
	}
}
