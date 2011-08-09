package br.com.plugitin.simpleexcelimport

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.util.CellReference
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator

import br.com.plugitin.simpleexcelimport.exception.InvalidValueException
import br.com.plugitin.simpleexcelimport.exception.NotADateColumnException

class SimpleExcelImport {

	/**
	 * Receives an excel file InputStream and a Sheet Configuration list.
	 * Sheet configuration should be as follows: 
	 * [
	 * 		name:"CDs",
	 * 		header:[
	 *			A:"Album Name",
	 *			B:"Artist",
	 *			C:"Year",
	 *			D:"Sold"
	 *			],
	 *		dateColumns:["Year"](optional),
	 *		startRow:2
	 *	]
	 * 
	 * @param excelInputStream
	 * @param sheetStructureList
	 * @return imported workbook object.
	 */
	def static excelImport(excelInputStream,sheetStructureList){
		if(excelInputStream){
			def workbook
			try{
				//Finds out the correct workbook version, 2003 or 2007.
				workbook = new WorkbookFactory().create(new PushbackInputStream(excelInputStream))
			}catch(all){
				throw new RuntimeException("Invalid File Type!")
			}
			importWorkbook(workbook,sheetStructureList)
		}
	}

	def static private importWorkbook(workbook,sheetStructureList){
		def workbookObject = [:]
		def evaluator

		//Get the appropriate evaluator based on the type of workbook provided by the factory.
		evaluator = workbook instanceof HSSFWorkbook? new HSSFFormulaEvaluator(workbook): new XSSFFormulaEvaluator(workbook)

		//Builds the workbook object based on the configuration provided, sheet by sheet.
		sheetStructureList?.each{sheetStructure->
			def sheetData = []
			def sheet = workbook.getSheet(sheetStructure.name)
			//Gathers data by row.
			def rowIterator = sheet?.rowIterator()
			if(rowIterator){
				loop_lines:
				for(def index = 0; rowIterator.hasNext(); index += 1) {
					def row = rowIterator.next()
					
					//Considers data from the established starting row forward.
					if(index >= sheetStructure.startRow-1){
						def rowData = [:]
						//Creates row data map.
						sheetStructure.header.each{entry->
							def cellContent = row.getCell(CellReference.convertColStringToIndex(entry.key))
							if(cellContent){
								try{
									rowData[entry.value] = resolveCell(cellContent,evaluator,sheetStructure.dateColumns?.contains(entry.value))
								}catch(all){
									throw new RuntimeException("Workbook contains error(s): '"+all.getMessage()+"'.")
								}
							}else{
								rowData[entry.value] = ""
							}
						}
						def emptyLine = true
						rowData.each { key, value ->
							if(!value.toString().isEmpty()) {
								emptyLine = false
							}
						}
						if(emptyLine) {	//stops the reading when the line is empty
							break loop_lines;
						}
						sheetData << rowData
					}
				}
			}
			workbookObject[sheetStructure.name] = sheetData
		}
		return workbookObject
	}

	def private static resolveCell(cellContent,evaluator,isDate){
		def returnValue
		switch(cellContent.getCellType()){
			case 0: //CELL_TYPE_NUMERIC
				returnValue = isDate?cellContent.getDateCellValue():cellContent.getNumericCellValue()
				break
			case 1://CELL_TYPE_STRING
				def content = cellContent.getStringCellValue()
				//When evaluating formulas a blank string may fall here, therefore the blank check
				returnValue = content?isDate?cellContent.getDateCellValue():content:""
				break
			case 2://CELL_TYPE_FORMULA
				returnValue = resolveCell(evaluator.evaluateInCell(cellContent),evaluator,isDate)
				break
			case 3://CELL_TYPE_BLANK
				returnValue = ""
				break
			case 4://CELL_TYPE_BOOLEAN
				returnValue = cellContent.getBooleanCellValue()
				break
			case 5://CELL_TYPE_ERROR
				throw new InvalidValueException()
			default:
				throw new RuntimeException("Unknown cell type!")
		}
		if(isDate && returnValue == null){
			throw new NotADateColumnException()
		}
		returnValue
	}
}
