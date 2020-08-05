;; # Excel Utilities
;; This file provides utility functions for reading `.xlsx` files.
;; It's a wrapper around a small part of the
;; [Apache POI project](http://poi.apache.org).
;; See the `incanter-excel` module from the
;; [Incanter](https://github.com/liebke/incanter) project for a more
;; thorough implementation.
;; TODO: Dates are not handled.

;; The function definitions progress from handling cells to rows, to sheets,
;; to workbooks.
(ns ontodev.excel
  (:require [clojure.tools.logging :as log]
            [clojure.string :as string]
            [clojure.java.io :as io])
  (:import
   (org.apache.poi.ss.usermodel Cell CellType Row Row$MissingCellPolicy
                                Sheet Workbook WorkbookFactory)))

;; ## Cells
;; I've found it hard to trust the Cell Type and Cell Style for data such as
;; integers. In this version of the code I'm converting each cell to STRING
;; type before reading it as a string and returning the string value.
;; This should be the literal value typed into the cell, except in the case
;; of formulae where it should be the result.
;; Conversion of the strings to other data types should be done as an
;; additional step.

(defn get-cell-string-value
  "Get the value of a cell as a string, by changing the cell type to 'string'
   and then changing it back."
  [cell]
  (let [ct    (.getCellType cell)
        cf    (if (= ct CellType/FORMULA) (.getCellFormula cell))
        _     (.setCellType cell CellType/STRING)
        value (.getStringCellValue cell)]
    (if (= ct CellType/FORMULA)
      (.setCellFormula cell cf)
      (.setCellType cell ct))
    value))

;; ## Rows
;; Rows are made up of cells. We consider the first row to be a header, and
;; translate its values into keywords. Then we return each subsequent row
;; as a map from keys to cell values.

(defn to-keyword
  "Take a string and return a properly formatted keyword."
  [s]
  (-> (or s "")
      string/trim
      string/lower-case
      (string/replace #"\s+" "-")
      keyword))

;; Note: it would make sense to use the iterator for the row. However that
;; iterator just skips blank cells! So instead we use an uglier approach with
;; a list comprehension. This relies on the workbook's setMissingCellPolicy
;; in `load-workbook`.
;; See `incanter-excel` and [http://stackoverflow.com/questions/4929646/how-to-get-an-excel-blank-cell-value-in-apache-poi]()

(defn read-row
  "Read all the cells in a row (including blanks) and return a list of values."
  [row]
  (for [i (range 0 (.getLastCellNum row))]
       (get-cell-string-value (.getCell row (.intValue i)))))

;; ## Sheets
;; Workbooks are made up of sheets, which are made up of rows.

(defn read-sheet
  "Given a workbook with an optional sheet name (default is 'Sheet1') and
   and optional header row number (default is '1') and optional function to parse headers
   return the data in the sheet as a vector of maps
   using the headers from the header row as the keys."
  ([workbook] (read-sheet workbook "Sheet1" 1))
  ([workbook sheet-name] (read-sheet workbook sheet-name 1))
  ([workbook sheet-name header-row]
   (read-sheet workbook sheet-name header-row to-keyword))
  ([workbook sheet-name header-row parse-header-fn]
   (log/debugf "Reading sheet '%s'" sheet-name)
   (let [sheet   (.getSheet workbook sheet-name)
         rows    (->> sheet (.iterator) iterator-seq (drop (dec header-row)))
         headers (map parse-header-fn (read-row (first rows)))
         data    (map read-row (rest rows))]
     (log/debugf "Read %d rows" (count rows))
     (vec (map (partial zipmap headers) data)))))

(defn list-sheets
  "Return a list of all sheet names."
  [workbook]
  (for [i (range (.getNumberOfSheets workbook))]
    (.getSheetName workbook i)))

(defn sheet-headers
  "Returns the headers (in their original forms, not as keywords) for a given sheet."
  [workbook sheet-name]
  (let [sheet (.getSheet workbook sheet-name)
        rows (->> sheet (.iterator) iterator-seq)]
    (read-row (first rows))))

;; ## Workbooks
;; An `.xlsx` file contains one workbook with one or more sheets.

(defn load-workbook
  "Load a workbook from a string path."
  [path]
  (log/debugf "Loading workbook:" path)
  (doto (WorkbookFactory/create (io/input-stream path))
        (.setMissingCellPolicy Row$MissingCellPolicy/CREATE_NULL_AS_BLANK)))


