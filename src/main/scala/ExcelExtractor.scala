import java.io.{File, FileInputStream, PrintWriter}

import org.apache.commons.lang3.StringUtils
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
  val NULL: String = "null value"

  def main(args: Array[String]): Unit = {
    val excelFileName = "Test1.xlsx"
    val excelFile = new FileInputStream(new File(excelFileName))

    val workbook = WorkbookFactory create excelFile

    val outFileName = "text.txt"
    val outfile = new PrintWriter(outFileName, "UTF-8")

    val sheet = workbook getSheetAt 0

    val rowEnd = sheet getLastRowNum

    val firstColumn = CellReference convertColStringToIndex "A"
    val lastColumn = CellReference convertColStringToIndex "E"

    var colAllListBuffer = new ListBuffer[String]

    for (rowNum <- 1 to rowEnd) {
      val row = sheet getRow rowNum
      for (colNum <- firstColumn to lastColumn) {
        val cell = row.getCell(colNum, Row.RETURN_NULL_AND_BLANK)
        colAllListBuffer += extractCellValue(cell)
      }
    }

    val columnValueList = colAllListBuffer.toList
    outfile.println("Excel 2 txt using Scala")

    columnValueList.foreach {
      outfile.println
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
          result = indexCell.getStringCellValue
          result = StringUtils.replace(result, "'", "''")
        case Cell.CELL_TYPE_NUMERIC =>
          result = StringUtils.removeEnd(String.valueOf(indexCell.getNumericCellValue), ".0")
        case Cell.CELL_TYPE_FORMULA =>
          val i: Int = indexCell.getCachedFormulaResultType
          if (i == Cell.CELL_TYPE_NUMERIC) {
            result = StringUtils.removeEnd(String.valueOf(indexCell.getNumericCellValue), ".0")
          }
          else if (i == Cell.CELL_TYPE_STRING) {
            result = indexCell.getStringCellValue
            result = StringUtils.replace(result, "'", "''")
          }
        case Cell.CELL_TYPE_BLANK =>
          result = NULL
        case _ =>
          result = NULL
      }
    }
    result = StringUtils upperCase result
    result
  }
}