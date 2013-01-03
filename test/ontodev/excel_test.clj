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
  (doall (map check-row data)))

