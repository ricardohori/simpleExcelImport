package br.com.plugitin.simpleexcelimport

import org.apache.commons.lang.StringUtils
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.util.CellReference
import org.apache.poi.ss.formula.eval.NotImplementedException
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator
import br.com.plugitin.simpleexcelimport.exception.*
import org.apache.poi.ss.usermodel.*

class SimpleExcelImport {
	
	/**
	 * Receives an excel file InputStream and a Sheet Configuration list.
	 * Sheet configuration should be as follows: 
	 * [
	 * 		name:"CDs",
	 * 		startRow:2		
	 * 		header:[
	 *			A:"Album Code",
	 *			B:"Artist",
	 *			C:"Year",
	 *			D:"Sold"
	 *			...
	 *			],
	 *      headerLine:[		(optional, tells the header row in which the tab is supposed to have the header names)
	 *      	row: 1,
	 *      	names: [columnNameA, columnNameB, ...]], 
	 *      ]
	 *      headerTypes:[		(optional, forces the type of each cell to be read. It can be informed only the desired headers instead of all of them)
	 *			A:CellType.NUMERIC,
	 *			B:CellType.STRING,
	 *			C:CellType.DATE,
	 *			...
	 *      ]
	 *	]
	 * 
	 * @param excelInputStream
	 * @param sheetStructureList
	 * @return imported workbook object.
	 */
	def static excelImport(excelInputStream,sheetStructureList){
		if(!excelInputStream){
			throw new IllegalArgumentException("The stream to the sheet must be specified for the import process")
		} else {
			Workbook workbook
			try{
				//Finds out the correct workbook version, 2003 or 2007.
				workbook = new WorkbookFactory().create(new PushbackInputStream(excelInputStream))
			}catch(all){
				throw new InvalidFileTypeException()
			}
			return importWorkbook(workbook,sheetStructureList)
		}
	}

	def static private importWorkbook(Workbook workbook,sheetStructureList){
		def workbookObject = [:]
		def evaluator

		//Get the appropriate evaluator based on the type of workbook provided by the factory.
		evaluator = workbook instanceof HSSFWorkbook? new HSSFFormulaEvaluator(workbook): new XSSFFormulaEvaluator(workbook)

		//Builds the workbook object based on the configuration provided, sheet by sheet.
		sheetStructureList?.each{sheetStructure->
			def sheetData = []
			Sheet sheet = workbook.getSheet(sheetStructure.name)
			if(!sheet){
				throw new TabNotFoundException(sheetStructure.name)
			}
			//Gathers data by row.
			def rowIterator = sheet?.rowIterator()
			if(rowIterator){
				loop_lines:
				for(def index = 0; rowIterator.hasNext(); index += 1) {
					Row row = rowIterator.next()
					def rowNumber = index+1
					
					//validate the presence of the column headers, if specified to do so
					if(sheetStructure.headerLine && rowNumber == sheetStructure.headerLine.row) {
						int idx = 0
						sheetStructure.headerLine.names.each{name->
							def cellContent = row.getCell(idx)
							def cellValue
							try{
								cellValue = cellContent?.getStringCellValue()
							}catch(all){}
							if(name != cellValue) {
								//throw an error when the column was not found on when its header differs from the expected one
								def columnLetter = CellReference.convertNumToColString(idx)
								throw new ColumnNotFoundException(tabName:sheetStructure.name, columnLetter:columnLetter, columnName:name, columnLine:rowNumber)
							}
							idx++
						}
					}
					
					//Considers data from the established starting row forward.
					if(rowNumber >= sheetStructure.startRow){
						def rowData = [:]
						//Creates row data map.
						sheetStructure.header.each{columnDef->
							def columnLetter = columnDef.key
							Cell cellContent = row.getCell(CellReference.convertColStringToIndex(columnLetter))
							if(cellContent){
								rowData[columnDef.value] = resolveCell(sheetStructure.name,rowNumber,columnLetter,cellContent,evaluator,sheetStructure.headerTypes ? sheetStructure.headerTypes[columnLetter] : null)
							}else{
								rowData[columnDef.value] = ""
							}
						}
						def emptyLine = true
						rowData.each { key, value ->
							if(StringUtils.isNotEmpty(value.toString())) {
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

	def private static resolveCell(tabName,rowNumber,columnLetter,Cell cellContent,evaluator,CellType desiredCellType){
		def returnValue
		def shouldBeDate = false
		def type = cellContent.getCellType()
		if(desiredCellType != null && type != Cell.CELL_TYPE_FORMULA && type != Cell.CELL_TYPE_ERROR && type != Cell.CELL_TYPE_BLANK){				
			if(desiredCellType == CellType.NUMERIC){
				type = Cell.CELL_TYPE_NUMERIC
			} else if(desiredCellType == CellType.STRING){
				type = Cell.CELL_TYPE_STRING
			} else if(desiredCellType == CellType.DATE){
				type = Cell.CELL_TYPE_NUMERIC
				shouldBeDate = true
			}
			cellContent.setCellType(type)
		}
		switch(type){
			case 0: //CELL_TYPE_NUMERIC
				returnValue = shouldBeDate ? cellContent.getDateCellValue() : cellContent.getNumericCellValue()
				break
			case 1://CELL_TYPE_STRING
				def content = cellContent.getStringCellValue()
				//When evaluating formulas a blank string may fall here, therefore the blank check
				returnValue = !content ? "" : content
				break
			case 2://CELL_TYPE_FORMULA
				try{
                    returnValue = resolveCell(tabName,rowNumber,columnLetter,evaluator.evaluateInCell(cellContent),evaluator,desiredCellType)
                }catch(NotImplementedException e){
                    throw new FormulaNotSupportedException(e, tabName, columnLetter, rowNumber)
                }
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
				throw new RuntimeException("Invalid cell type at ${rowNumber}, column ${columnLetter} within tab ${tabName}")
		}
		returnValue
	}
}
