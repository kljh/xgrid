import * as rp from "./excel_range_parse.js"
import * as xlf from "./excel_formula_transform.js"
import * as jsf from "./js_formula_tranform.js"
import * as shx from "./sheet_exec.js"

rp.parse_range_bijection_test();
xlf.parse_and_transfrom_test();
jsf.parse_and_transfrom_test();
shx.spreadsheet_exec_test();

