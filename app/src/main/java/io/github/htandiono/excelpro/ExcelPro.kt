package io.github.htandiono.excelpro

import android.content.Context
import android.net.Uri
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.io.InputStream
import java.io.OutputStream

/**
 * A utility class to read and write data from/to Excel files in Android.
 * Supports both .xls (HSSFWorkbook) and .xlsx (XSSFWorkbook) formats.
 *
 * @param context The application context.
 */
class ExcelPro(private val context: Context) {

    private var workbook: Workbook? = null
    private var sheet: Sheet? = null
    private var fileUri: Uri? = null
    private lateinit var headerMap: Map<String, Int>

    companion object {
        const val PNO_COLUMN_HEADER = "PNO"
    }

    /**
     * Opens an Excel file from a given Uri and prepares it for reading.
     * It autodetects the file format (.xls or .xlsx).
     *
     * @param uri The Uri of the Excel file to open.
     * @throws Exception if the file cannot be opened or is not a valid Excel file.
     */
    fun openFile(uri: Uri) {
        this.fileUri = uri
        val inputStream: InputStream = context.contentResolver.openInputStream(uri)
            ?: throw Exception("Could not open input stream from Uri.")

        inputStream.use { stream ->
            workbook = try {
                // Try reading as .xlsx first
                XSSFWorkbook(stream)
            } catch (e: Exception) {
                // Fallback to .xls if it fails
                // We need a fresh stream, so let's reopen it.
                context.contentResolver.openInputStream(uri)?.let { HSSFWorkbook(it) }
            }
        }

        if (workbook == null) {
            throw Exception("Unsupported file format or corrupted file.")
        }

        // We'll work with the first sheet by default
        sheet = workbook?.getSheetAt(0)
        mapHeaderColumns()
    }

    /**
     * Reads all data from the opened Excel sheet.
     *
     * @return A list of maps, where each map represents a row (keyed by header).
     * @throws IllegalStateException if no file is opened.
     */
    fun readData(): List<Map<String, String>> {
        val currentSheet = sheet ?: throw IllegalStateException("No Excel sheet is open.")
        if (!::headerMap.isInitialized) throw IllegalStateException("Header row not mapped.")

        val data = mutableListOf<Map<String, String>>()
        val headerKeys = headerMap.keys.toList()

        // Start from the first data row (skip header)
        for (i in 1..currentSheet.lastRowNum) {
            val row = currentSheet.getRow(i) ?: continue
            val rowData = mutableMapOf<String, String>()
            headerKeys.forEach { headerName ->
                val cellIndex = headerMap[headerName]!!
                val cell = row.getCell(cellIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)
                rowData[headerName] = getCellStringValue(cell)
            }
            data.add(rowData)
        }
        return data
    }

    /**
     * Finds a specific row by its PNO value.
     *
     * @param pno The Palm Number (PNO) to search for.
     * @return A map representing the row data if found, otherwise null.
     */
    fun findRowByPno(pno: String): Map<String, String>? {
        val pnoColumnIndex = headerMap[PNO_COLUMN_HEADER]
            ?: return null // PNO column doesn't exist

        // Start from the first data row (skip header)
        for (i in 1..sheet!!.lastRowNum) {
            val row = sheet!!.getRow(i) ?: continue
            val cell = row.getCell(pnoColumnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)
            if (cell != null && getCellStringValue(cell) == pno) {
                return mapRowToHeader(row)
            }
        }
        return null
    }

    /**
     * Updates a specific cell in the sheet based on PNO and a column header.
     * If the column doesn't exist, it will be created.
     *
     * @param pno The PNO of the row to update.
     * @param columnHeader The header of the column to update.
     * @param value The new value to write into the cell.
     * @return True if the update was successful, false otherwise.
     */
    fun updateCell(pno: String, columnHeader: String, value: String): Boolean {
        val currentSheet = sheet ?: throw IllegalStateException("No Excel sheet is open.")
        val pnoColumnIndex = headerMap[PNO_COLUMN_HEADER]
            ?: return false // PNO column must exist

        // Find the row index for the given PNO
        var rowIndex = -1
        for (i in 1..currentSheet.lastRowNum) {
            val row = currentSheet.getRow(i) ?: continue
            val cell = row.getCell(pnoColumnIndex)
            if (cell != null && getCellStringValue(cell) == pno) {
                rowIndex = i
                break
            }
        }
        if (rowIndex == -1) return false // PNO not found

        // Find or create the column index for the given header
        var columnIndex = headerMap[columnHeader]
        if (columnIndex == null) {
            // Column doesn't exist, so create it
            val headerRow = currentSheet.getRow(0)
            val newCellIndex = headerRow.lastCellNum.toInt()
            headerRow.createCell(newCellIndex).setCellValue(columnHeader)
            // Re-map headers
            mapHeaderColumns()
            columnIndex = headerMap[columnHeader]!!
        }

        // Get the row and update the cell
        val row = currentSheet.getRow(rowIndex) ?: currentSheet.createRow(rowIndex)
        row.createCell(columnIndex).setCellValue(value)

        return true
    }

    /**
     * Saves the changes back to the original file Uri.
     *
     * @throws IllegalStateException if no file is open.
     * @throws Exception if the file cannot be written to.
     */
    fun saveFile() {
        val currentWorkbook = workbook ?: throw IllegalStateException("Workbook is not open.")
        val uri = fileUri ?: throw IllegalStateException("File Uri is not set.")

        try {
            val outputStream: OutputStream = context.contentResolver.openOutputStream(uri, "w")
                ?: throw Exception("Could not open output stream for writing.")
            outputStream.use { stream ->
                currentWorkbook.write(stream)
            }
        } catch (e: Exception) {
            // Propagate the exception to be handled by the UI
            throw Exception("Failed to save Excel file: ${e.message}")
        }
    }

    /**
     * Closes the workbook to release resources.
     */
    fun close() {
        workbook?.close()
        workbook = null
        sheet = null
        fileUri = null
    }

    // --- Private Helper Methods ---

    private fun mapHeaderColumns() {
        val currentSheet = sheet ?: return
        val headerRow = currentSheet.getRow(0) ?: return
        val tempHeaderMap = mutableMapOf<String, Int>()
        for (cell in headerRow) {
            tempHeaderMap[getCellStringValue(cell)] = cell.columnIndex
        }
        headerMap = tempHeaderMap.toMap()
    }

    private fun mapRowToHeader(row: Row): Map<String, String> {
        val rowData = mutableMapOf<String, String>()
        headerMap.forEach { (header, index) ->
            val cell = row.getCell(index, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)
            rowData[header] = getCellStringValue(cell)
        }
        return rowData
    }

    private fun getCellStringValue(cell: Cell?): String {
        return when (cell?.cellType) {
            CellType.STRING -> cell.stringCellValue
            CellType.NUMERIC -> {
                // Check if it's a date and format it, otherwise get numeric value
                if (DateUtil.isCellDateFormatted(cell)) {
                    cell.dateCellValue.toString()
                } else {
                    // Avoid ".0" for whole numbers
                    val number = cell.numericCellValue
                    if (number == number.toLong().toDouble()) {
                        number.toLong().toString()
                    } else {
                        number.toString()
                    }
                }
            }
            CellType.BOOLEAN -> cell.booleanCellValue.toString()
            CellType.FORMULA -> {
                try {
                    // Evaluate formula and get the result as a string
                    val evaluator = workbook?.creationHelper?.createFormulaEvaluator()
                    val cellValue = evaluator?.evaluate(cell)
                    cellValue?.formatAsString() ?: ""
                } catch (e: Exception) {
                    "FORMULA_ERROR"
                }
            }
            else -> ""
        }.trim()
    }
}