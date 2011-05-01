package br.com.futuresolutions.simpleexcelimport

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.util.CellReference
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator

import br.com.futuresolutions.simpleexcelimport.exception.InvalidValueException
import br.com.futuresolutions.simpleexcelimport.exception.NotADateColumnException

class SimpleExcelImport {
	
	def static excelImport(excelInputStream,sheetStructureList){
		if(excelInputStream){
			def workbook
			try{
				workbook = new WorkbookFactory().create(new PushbackInputStream(excelInputStream))
			}catch(all){
				throw new RuntimeException("Invalid File Type!")
			}
			importExcel(workbook,sheetStructureList)	
		}
	}
	
	def static private importExcel(workbook,sheetStructureList){
		def workbookObject = [:]
		def evaluator
		
		if(workbook instanceof HSSFWorkbook){
			evaluator = new HSSFFormulaEvaluator(workbook)
		}else{
			evaluator = new XSSFFormulaEvaluator(workbook)
		}

		sheetStructureList?.each{sheetStructure->
			def sheetValuesList = []
			def sheet = workbook.getSheet(sheetStructure.name)
			sheet?.rowIterator().eachWithIndex{row,index->
				if(index >= sheetStructure.startRow-1){
					def sheetMap = [:]
					sheetStructure.header.each{entry->
						def cellContent = row.getCell(CellReference.convertColStringToIndex(entry.key))
						if(cellContent){
							sheetMap[entry.value] = resolveCell(cellContent,evaluator,sheetStructure.dateColumns?.contains(entry.value))
						}else{
							sheetMap[entry.value] = ""
						}
					}
					sheetValuesList << sheetMap
				}
			}
			workbookObject[sheetStructure.name] = sheetValuesList			
		}
		return workbookObject
	}
	
	def private static resolveCell(cellContent,evaluator,isDate){
		def returnValue
		try{
			switch(cellContent.getCellType()){
				case 0: //CELL_TYPE_NUMERIC
					returnValue = isDate?cellContent.getDateCellValue():cellContent.getNumericCellValue()
					break
				case 1://CELL_TYPE_STRING
					returnValue = isDate?cellContent.getDateCellValue():cellContent.getStringCellValue()
					break
				case 2://CELL_TYPE_FORMULA
					returnValue = resolveCell(evaluator.evaluateInCell(cellContent),evaluator,isDate)
					break
				case 3://CELL_TYPE_BLANK
					returnValue = ""
					break
				case 4://CELL_TYPE_BOOLEAN
					returnValue = isDate?cellContent.getDateCellValue():cellContent.getBooleanCellValue()
					break
				case 5://CELL_TYPE_ERROR
					throw new InvalidValueException()
				default:
					throw new RuntimeException("Unknown cell type!")
			}
			if(isDate && returnValue == null){
				throw new NotADateColumnException()
			}
		}catch(InvalidValueException e){
			throw new RuntimeException(e.getMessage())
		}catch(NotADateColumnException e){
			throw new RuntimeException(e.getMessage())
		}catch(RuntimeException re){
			throw new RuntimeException("Workbook contains error(s): '"+re.getMessage()+"'.")
		}
		returnValue
	}
}
