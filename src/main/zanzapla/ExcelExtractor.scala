package zanzapla

import java.io.{File, FileInputStream, PrintWriter}
import java.time.LocalDate
import org.apache.poi.ss.usermodel.{Cell, Row, WorkbookFactory}
import org.apache.poi.ss.util.CellReference

import scala.collection.mutable.ListBuffer

class ExcelExtractor {

}

/**
  * Created by @Zanza00 on 18/09/2015.
  *
  * A simple excel extractor, made in Scala
  */


object excel {
  val NULL: String = "NULL"
  val charsToReject = "\t'".toSet

  def main(args: Array[String]): Unit = {
    val excelFileName = "Test1.xlsx"
    val excelFile = new FileInputStream(new File(excelFileName))

    val workbook = WorkbookFactory create excelFile

    val outFileName = "output.sql"
    val outfile = new PrintWriter(outFileName, "UTF-8")

    val sheet = workbook getSheetAt 0

    val rowEnd = sheet getLastRowNum

    val firstColumn = CellReference convertColStringToIndex "A"
    val lastColumn = CellReference convertColStringToIndex "E"


    outfile.println("--Excel 2 txt using Scala")
    outfile.println("--Excel file: " + excelFileName)
    outfile.println("--Executed on: " + LocalDate.now())

    for (rowNum <- 1 to rowEnd) {
      var colAllListBuffer = new ListBuffer[String]
      colAllListBuffer += "INSERT INTO TABLE_NAME (NUMBERS, BIG_NUMBERS, DECIMALS, STRINGS, LANGUAGES, FIXED_VALUE) VALUES("
      val row = sheet getRow rowNum
      for (colNum <- firstColumn to lastColumn) {
        val cell = row.getCell(colNum, Row.RETURN_NULL_AND_BLANK)
        colAllListBuffer += extractCellValue(cell)

        colAllListBuffer += ","
      }
      colAllListBuffer += "'FIXED')"

      val columnValueList = colAllListBuffer.toList
      columnValueList.foreach {
        outfile.print
      }

      outfile.println(";")

    }


    //close every file
    excelFile.close()
    outfile.close()


  }

  private def extractCellValue(indexCell: Cell): String = {
    var result: String = NULL
    if (indexCell != null) {
      indexCell.getCellType match {
        case Cell.CELL_TYPE_STRING =>
          result = indexCell.getStringCellValue filterNot charsToReject
          result = "'" + result + "'"
        case Cell.CELL_TYPE_NUMERIC =>
          result = String.valueOf(indexCell.getNumericCellValue)
        case Cell.CELL_TYPE_FORMULA =>
          indexCell.getCachedFormulaResultType match {
            case Cell.CELL_TYPE_NUMERIC =>
              result = String.valueOf(indexCell.getNumericCellValue)
            case Cell.CELL_TYPE_STRING =>
              result = indexCell.getStringCellValue filterNot charsToReject
              result = "'" + result + "'"
            case _ =>
              result = NULL
          }
        case Cell.CELL_TYPE_BLANK =>
          result = NULL
        case _ =>
          result = NULL
      }
    }
    result.toUpperCase
  }
}