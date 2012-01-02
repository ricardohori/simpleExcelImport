package br.com.plugitin.simpleexcelimport.exception

class InvalidFileTypeException extends RuntimeException {
	
	@Override
	public String getMessage() {
		return "Invalid file type!"
	}
}
