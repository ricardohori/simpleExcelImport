package br.com.plugitin.exception

class InvalidValueException extends Exception {
	@Override
	public String getMessage() {
		return "Invalid value(s)!"
	}
}
