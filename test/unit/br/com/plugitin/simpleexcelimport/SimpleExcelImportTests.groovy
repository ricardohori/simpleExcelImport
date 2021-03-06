package br.com.plugitin.simpleexcelimport

import grails.test.*

import java.text.SimpleDateFormat

import br.com.plugitin.simpleexcelimport.exception.ColumnNotFoundException
import br.com.plugitin.simpleexcelimport.exception.InvalidValueException
import br.com.plugitin.simpleexcelimport.exception.TabNotFoundException
import br.com.plugitin.simpleexcelimport.exception.FormulaNotSupportedException

class SimpleExcelImportTests extends GrailsUnitTestCase {

    void testExcelImport2003() {
		def excelFile = getInputStream("test/resources/test2003.xls")
		testWorkbook excelFile
    }
	
	void testExcelImport2007() {
		def excelFile = getInputStream("test/resources/test2007.xlsx")
		testWorkbook excelFile
	}
	
	void testExcelImportNameException() {
		def excelFile = getInputStream("test/resources/testNameError.xlsx")
		try{
			testWorkbook excelFile
			fail()
		}catch(RuntimeException re){}
	}

	void testExcelImportValueException() {
		def excelFile = getInputStream("test/resources/testValueError.xlsx")
		try{
			testWorkbook excelFile
			fail()
		}catch(Exception re){
			assertEquals InvalidValueException.class, re.getClass()
		}
	}
	
	void testNotAWorkbookException() {
		def excelFile = getInputStream("test/resources/notAWorkbook.txt")
		try{
			testWorkbook excelFile
		}catch(RuntimeException re){
			assertEquals "Invalid file type!", re.getMessage()
		}
	}
	
	void testReadUntilBlankLine() {
		def excelFile = getInputStream("test/resources/readUntilBlankLine.xlsx")
		def workbook = SimpleExcelImport.excelImport(excelFile, [styleSheetBooks()])
		assertEquals 3, workbook.Books.size()
	}
	
	/**
	 * Asserts that each expected tab must be found into the sheet
	 */
	void testErrorTabNotFound() {
		def excelFile = getInputStream("test/resources/blank.xls")
		try {						
			SimpleExcelImport.excelImport(excelFile, [styleSheetBooks()])
			fail("An error was expected")
		} catch (Exception e) {
			if(!(e instanceof TabNotFoundException)){
				fail("TabNotFoundException was expected")
			}
		}
	}
	
    /**
 	 * Asserts that the column headers must be found in the headerRow whenever the headerRow is set to a certain tab to be read
	 */
	void testErrorColumnNotFound() {
		def structure = [:]
		structure.putAll(styleSheetBooks())
		structure.headerLine = [row:1, names:["Book Name", "Author", "Year"]]
		
		def excelFile = getInputStream("test/resources/testColumnNotFound.xls")
		try {
			SimpleExcelImport.excelImport(excelFile, [structure])
			fail("An error was expected")
		} catch (Exception e) {
			if(e instanceof ColumnNotFoundException){
				assertEquals "Books", e.tabName
				assertEquals "B", e.columnLetter
				assertEquals "Author", e.columnName
				assertEquals 1, e.columnLine
			} else {
				assertEquals ColumnNotFoundException.class, e.getClass()
			}
		}
	}

    void testErrorFormulaNotSupported() {
        def structure = [:]
        structure.putAll(styleSheetBooks())
        structure.headerLine = [row:1, names:["Book Name", "Author", "Year"]]

        def excelFile = getInputStream("test/resources/testFormulaNotSupported.xls")
        try {
            SimpleExcelImport.excelImport(excelFile, [structure])
            fail("An error was expected")
        } catch (Exception e) {
            if(e instanceof FormulaNotSupportedException){
                assertEquals "Books", e.tabName
                assertEquals "B", e.columnLetter
                assertEquals 2, e.columnLine
            } else {
                assertEquals FormulaNotSupportedException.class, e.getClass()
            }
        }
    }

	private void testWorkbook(excelFile){
		def sheetStructureList = []
		def testSheet = [
			name:"CDs",
			startRow:2,
			header:[
				A:"Album Name",
				B:"Artist",
				C:"Year",
				D:"Sold"
				],
			headerTypes:[
				C:CellType.DATE,
			],
		]
		sheetStructureList << testSheet
		sheetStructureList << styleSheetBooks()
		
		def workbook = SimpleExcelImport.excelImport(excelFile, sheetStructureList)
		
		assertNotNull "CDs sheet shouldn't be null!",workbook.CDs
		assertNotNull "Books sheet shouldn't be null!",workbook.Books
		
		def cds = workbook.CDs
		
		assertTrue "Album Name should be a String!",cds[0]["Album Name"] instanceof String
		assertEquals "First album should be Rubber Soul!","Rubber Soul",cds[0]["Album Name"]
		assertEquals "Second album should be Revolver!","Revolver",cds[1]["Album Name"]
		assertEquals "Third album should be Sgt. Pepper's Lonely Hearts Club Band!","Sgt. Pepper's Lonely Hearts Club Band",cds[2]["Album Name"]
		assertEquals "Fourth album should be Magical Mystery Tour!","Magical Mystery Tour",cds[3]["Album Name"]
		assertEquals "Fifth album should be White Album!","White Album",cds[4]["Album Name"]
		
		assertTrue "Artist should be a String!",cds[0]["Artist"] instanceof String
		assertEquals "First artist should be The Beatles!","The Beatles",cds[0]["Artist"]
		assertEquals "Second artist should be The Beatles!","The Beatles",cds[1]["Artist"]
		assertEquals "Third artist should be The Beatles!","The Beatles",cds[2]["Artist"]
		assertEquals "Fourth artist should be The Beatles!","The Beatles",cds[3]["Artist"]
		assertEquals "Fifth artist should be The Beatles!","The Beatles",cds[4]["Artist"]
		
		def sdf = new SimpleDateFormat("dd/MM/yyyy")
		assertTrue "Artist should be a Date!",cds[0]["Year"] instanceof Date
		assertEquals "First year should be 18/05/1965!",sdf.parse("18/05/1965"),cds[0]["Year"]
		assertEquals "Second year should be 18/05/1966!",sdf.parse("19/05/1966"),cds[1]["Year"]
		assertEquals "Third year should be 18/05/1967!",sdf.parse("20/05/1967"),cds[2]["Year"]
		assertEquals "Fourth year should be 18/05/1967!",sdf.parse("20/05/1967"),cds[3]["Year"]
		assertEquals "Fifth year should be 18/05/1968!",sdf.parse("21/05/1968"),cds[4]["Year"]
		
		assertTrue "Sold should be a Numeric value!",cds[0]["Sold"] instanceof Number
		assertEquals "First sold should be 23880!",23880,cds[0]["Sold"]
		assertEquals "Second sold should be 24246!",24246,cds[1]["Sold"]
		assertEquals "Third sold should be 24612!",24612,cds[2]["Sold"]
		assertEquals "Fourth sold should be 24612!",24612,cds[3]["Sold"]
		assertEquals "Fifth sold should be 24979!",24979,cds[4]["Sold"]
		
		def books = workbook.Books
		
		assertTrue "Book Name should be a String!",books[0]["Book Name"] instanceof String
		assertEquals "First book name should be The Gates of Rome!","The Gates of Rome",books[0]["Book Name"]
		assertEquals "Second book name should be The Death of Kings!","The Death of Kings",books[1]["Book Name"]
		assertEquals "Third book name should be The Field of Swords!","The Field of Swords",books[2]["Book Name"]
		assertEquals "Fourth book name should be The Gods of War!","The Gods of War",books[3]["Book Name"]
		
		assertTrue "Author should be a String!",books[0]["Author"] instanceof String
		assertEquals "First author should be Conn Iggulden!","Conn Iggulden",books[0]["Author"]
		assertEquals "Second author should be Conn Iggulden!","Conn Iggulden",books[1]["Author"]
		assertEquals "Third author should be Conn Iggulden!","Conn Iggulden",books[2]["Author"]
		assertEquals "Fourth author should be Conn Iggulden!","Conn Iggulden",books[3]["Author"]
		
		assertTrue "Year should be a Date!",books[0]["Year"] instanceof Date
		assertEquals "First year should be 25/06/2005!",sdf.parse("25/06/2005"),books[0]["Year"]
		assertEquals "Second year should be 26/06/2006!",sdf.parse("26/06/2006"),books[1]["Year"]
		assertEquals "Third year should be be blank!","",books[2]["Year"]
		assertEquals "Fourth year should be 28/06/2008!",sdf.parse("28/06/2008"),books[3]["Year"]
	}
	
	void testForcedTypeReading(){
		def testSheet = [
			name:"Books",
			startRow:2,
			header:[
				A:"Book code",
				B:"Book Name",
			],
			headerTypes:[
				A:CellType.NUMERIC,
				B:CellType.STRING
			],
		]
		def excelFile = getInputStream("test/resources/testForcedReadingTypes.xlsx")
		def workbook = SimpleExcelImport.excelImport(excelFile, [testSheet])
		def books = workbook.Books
		assertEquals 5, books[0]["Book code"]
		assertEquals "10", books[0]["Book Name"]
		assertEquals 6, books[1]["Book code"]
		assertEquals "20", books[1]["Book Name"]
	}

    private def styleSheetBooks() {
        return [
            name:"Books",
            header:[
                A:"Book Name",
                B:"Author",
                C:"Year"
            ],
            headerTypes:[
                    C:CellType.DATE,
            ],
            startRow:2
        ]
    }

	private def getInputStream(filename){
		new FileInputStream(new File(filename))
	}
}