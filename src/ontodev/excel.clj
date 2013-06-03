;; # Excel Utilities
;; This file provides utility functions for reading `.xlsx` files.
;; It's a wrapper around a small part of the 
;; [Apache POI project](http://poi.apache.org).
;; See the `incanter-excel` module from the
;; [Incanter](https://github.com/liebke/incanter) project for more.
;;
;; The functions build from handling cells to rows, to sheets, to workbooks.
(ns ontodev.excel
  (:require [clojure.tools.logging :as log]
            [clojure.string :as string]
            [clojure.java.io :as io])
  (:import
    (org.apache.poi.ss.usermodel Cell Row Sheet Workbook DateUtil WorkbookFactory)
    (org.apache.poi.xssf.usermodel XSSFWorkbook)))

;; ## Cells
;; This is the trickiest piece of the code -- it is not complete, nor is it
;; well-tested. You can specify cell types in Excel, but these settings are
;; often ignored, for instance in the case of a text cell containing just
;; digits.
;; We define a multimethod that switches on the cell type, then methods for
;; each cell type.
;; See [https://github.com/liebke/incanter/blob/master/modules/incanter-excel/src/incanter/excel]()

(defmulti get-cell-value
  "Get the cell value depending on the cell type. Note that numeric cells
   can also contain dates, so we handle this special case."
  (fn [cell]
      (when (not= nil cell)
        (let [ct (. cell getCellType)]
          (if (not (= Cell/CELL_TYPE_NUMERIC ct))
            ct
            (if (DateUtil/isCellDateFormatted cell)
              :date
              ct))))))

(defmethod get-cell-value Cell/CELL_TYPE_BLANK   [cell] nil)

;; This is a partial implementation that expects a numeric value.
;; TODO: Handle cell types other than numeric.
(defmethod get-cell-value Cell/CELL_TYPE_FORMULA [cell]
  (let [val (.
             (.. cell
                 getSheet
                 getWorkbook
                 getCreationHelper
                 createFormulaEvaluator)
             evaluate cell)
        evaluated-type (. val getCellType)]
    (if (= 1 (.getDataFormat (.getCellStyle cell)))
      (.intValue (. val getNumberValue))  
      (. val getNumberValue))))

(defmethod get-cell-value Cell/CELL_TYPE_BOOLEAN [cell]
  (. cell getBooleanCellValue))

(defmethod get-cell-value Cell/CELL_TYPE_STRING  [cell]
  (. cell getStringCellValue))

;; Returns the value of a cell as a number, as far as possible. Handles two 
;; special cases based on the CellStyle: integers, and cells that were
;; specified as strings but then called numeric. The default result is a
;; double.
;; TODO: This implementation is incomplete.
(defmethod get-cell-value Cell/CELL_TYPE_NUMERIC [cell]
  (case (.getDataFormat (.getCellStyle cell))
    1  (.intValue (.getNumericCellValue cell)) ; Integer
    49 (str (.intValue (.getNumericCellValue cell))) ; Integer as string TODO: this might be a bad idea
    (.getNumericCellValue cell))) ; default: Double

(defmethod get-cell-value :date [cell]
  (. cell getDateCellValue))

(defmethod get-cell-value :default [cell]
  (str "Unknown cell type " (. cell getCellType)))


;; ## Rows
;; Rows are made up of cells. We consider the first row to be a header, and 
;; translate its values into keywords. Then we return each subsequent row
;; as a map from keys to cell values.

(defn to-keyword
  "Take a string and return aa properly formatted keyword."
  [s]
  (-> s
      string/trim
      string/lower-case
      (string/replace #"\s+" "-")
      keyword))

;; Note that the row iterator just skips blank cells, so instead we use an
;; uglier approach with a list comprehension. This relies on the workbook's
;; setMissingCellPolicy above.
;; See `incanter-excel` and [http://stackoverflow.com/questions/4929646/how-to-get-an-excel-blank-cell-value-in-apache-poi]()
(defn read-row
  "Read all the cells in a row (including blanks) and return a list of values."
  [row]
  (for [i (range (.getFirstCellNum row) (.getLastCellNum row))]
       (get-cell-value (.getCell row (.intValue i)))))

;; ## Sheets
;; Workbooks are made up of sheets, which are made up of rows.

(defn read-sheet
  "Read a sheet from a workbook and return the data as a vector of maps."
  ([workbook sheet-name] (read-sheet workbook sheet-name 1))
  ([workbook sheet-name header-row] 
   (log/debugf "Reading sheet '%s'" sheet-name)
   (let [sheet   (.getSheet workbook sheet-name)
         rows    (drop (- header-row 1) (iterator-seq (. sheet iterator)))
         headers (map to-keyword (read-row (first rows))) 
         data    (map read-row (rest rows))]
     (log/debugf "Read %d rows" (count rows))
     (vec (map (partial zipmap headers) data)))))

(defn read-sheet-simple
  "Read a sheet from a workbook as rows"
  [workbook sheet-name]
  (log/debugf "Reading sheet '%s'" sheet-name)
  (let [sheet   (.getSheet workbook sheet-name)
        rows    (iterator-seq (. sheet iterator))]
    (vec rows)))

;; ## Workbooks
;; An `.xlsx` file contains one workbook with one or more sheets.

(defn load-workbook
  "Load a workbook from a string path."
  [path]
  (log/info "Loading workbook:" path)
  (doto (WorkbookFactory/create (io/input-stream path))
        (.setMissingCellPolicy Row/CREATE_NULL_AS_BLANK)))
