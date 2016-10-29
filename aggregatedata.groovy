@Grapes([
        @Grab('org.apache.poi:poi:3.10.1'),
        @Grab('org.apache.poi:poi-ooxml:3.10.1')])
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import static org.apache.poi.ss.usermodel.Cell.*

import java.nio.file.Paths
import groovy.json.JsonOutput

def header = []
def values = []

