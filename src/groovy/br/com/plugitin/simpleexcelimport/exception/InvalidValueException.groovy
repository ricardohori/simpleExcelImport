package br.com.plugitin.simpleexcelimport.exception

class InvalidValueException extends Exception {
	@Override
	public String getMessage() {
		return "Invalid value(s)!"
	}
}
