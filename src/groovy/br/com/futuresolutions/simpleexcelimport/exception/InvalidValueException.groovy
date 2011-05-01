package br.com.futuresolutions.simpleexcelimport.exception

class InvalidValueException extends Exception {
	@Override
	public String getMessage() {
		return "Invalid value(s)!"
	}
}
