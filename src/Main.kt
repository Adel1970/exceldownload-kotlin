import org.apache.poi.xssf.streaming.SXSSFWorkbook
import java.io.FileOutputStream
import java.util.*

fun main(args: Array<String>) {
        createExcel()
}

fun createExcel() {

        val workbook = SXSSFWorkbook()
        val sheet = workbook.createSheet("Sell Report")
        val data = generateTestData()
        val headerRow = sheet.createRow(0)
        for ((index, title)  in data.keys.withIndex()) {
           headerRow.createCell(index).setCellValue(title)
        }
        val columns = data.values
        val numberOfRows = data.values.elementAt(0).size

        for (i in 1 until numberOfRows) {
           val row = sheet.createRow(i)
            for(j in columns.indices) {
                row.createCell(j).setCellValue(columns.elementAt(j).elementAt(i))
            }
        }

        val outputStream = FileOutputStream("kotlin_excel.xlsx")
        workbook.write(outputStream)
}

fun generateTestData(): TreeMap<String, MutableList<String>> {

    val data = TreeMap<String, MutableList<String>>()
    for(i in 1..70) {
        for (j in 1..700000)
            if (j == 1) {
                data["title_$i"] = MutableList(j) {"testval_$j"}
            } else {
                data["title_$i"]!!.add("testval_$j")
            }
    }
    return data
}