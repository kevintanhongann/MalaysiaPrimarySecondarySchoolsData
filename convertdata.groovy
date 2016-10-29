@Grapes([
        @Grab('org.apache.poi:poi:3.10.1'),
        @Grab('org.apache.poi:poi-ooxml:3.10.1')])
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import static org.apache.poi.ss.usermodel.Cell.*

import java.nio.file.Paths
import groovy.json.JsonOutput

def header = []
def values = []

Paths.get('sekolahrendahdanmenengahmac2015.xlsx').withInputStream { input ->

    def workbook = new XSSFWorkbook(input)
    def sheet = workbook.getSheetAt(0)

    //header is obtained at row 1
    for (cell in sheet.getRow(1).cellIterator()) {
        header << cell.stringCellValue
        println header
    }

    def headerFlag = true

    sheet.rowIterator().eachWithIndex { val, index ->
        println 'index ' + index
        if (headerFlag) {
            headerFlag = false
        }
        def rowData = [:]
        if (index != 0 && index != 1) {
            //val is row
            val.cellIterator().eachWithIndex { cellValue, callIndex ->
                //callIndex is 1 to 10
                if (callIndex != 0) {
                    def value = ''
                    switch (cellValue.cellType) {
                        case CELL_TYPE_STRING:
                            value = cellValue.stringCellValue
                            break
                        case CELL_TYPE_STRING:
                            value = cellValue.numericCellValue
                            break
                        default:
                            value = ''
                    }
                    if (cellValue.columnIndex != 0) {
                        rowData << ["${header[cellValue.columnIndex]}": value]
                    }
                }
            }
            values << rowData
        }
    }
}

Paths.get('convertedSchoolsData.xlsx.json').withWriter { jsonWriter ->
    jsonWriter.write JsonOutput.prettyPrint(JsonOutput.toJson(values))
}
