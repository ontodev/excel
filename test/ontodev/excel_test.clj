(ns ontodev.excel-test
  (:use midje.sweet
        ontodev.excel))

(defn check-row
  [row]
  (facts "check-row"
    (:integer row) => "1001" 
    (:float row) => "1001.01" 
    (:formula row) => "2002.01"))

(let [workbook (load-workbook "resources/test.xlsx")
      data     (read-sheet workbook)]
  (doall (map check-row data))
  (fact "sheet names" (list-sheets workbook) => (just ["Sheet1" "Foo" "Bar"]))
  (fact "sheet headers" (sheet-headers workbook "Sheet1") => (just ["Format" "Integer" "Float" "Formula"])))
