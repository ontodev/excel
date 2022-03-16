(ns ontodev.excel-test
  (:require [clojure.test :refer :all])
  (:use midje.sweet
        ontodev.excel))

(deftest test-to-keyword-takes-nil
  (is (nil? (to-keyword nil))))

(deftest test-to-keyword-valid
  (is (= :keyword-one (to-keyword "keyword one")))
  (is (= :keyword-one (to-keyword "Keyword One")))
  (is (= :keyword-one (to-keyword "   KeyWord oNe   "))))

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
