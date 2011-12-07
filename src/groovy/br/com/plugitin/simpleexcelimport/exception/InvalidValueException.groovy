package br.com.plugitin.simpleexcelimport.exception

class InvalidValueException extends RuntimeException {
	@Override
	public String getMessage() {
		return "Invalid value(s)!"
	}
}
