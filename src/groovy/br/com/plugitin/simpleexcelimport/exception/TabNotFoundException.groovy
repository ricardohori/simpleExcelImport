package br.com.plugitin.simpleexcelimport.exception

class TabNotFoundException extends RuntimeException {
	
	private def tabName
	
	public TabNotFoundException() {
		this("")
	}
	
	public TabNotFoundException(def tabName){
		this.tabName = tabName ?: ""
	}
	
	@Override
	public String getMessage() {
		return "Expected tab '${tabName}' was not found within the sheet"
	}
}
