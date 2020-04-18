# ontodev.excel

A thin [Clojure](http://clojure.org) wrapper around a small part of [Apache POI](http://poi.apache.org) for reading `.xlsx` files. 

For a more complete implementation, see the `incanter-excel` module from the [Incanter](https://github.com/liebke/incanter) project.

## Usage

Add `ontodev/excel` to your [Leiningen](http://leiningen.org/) project dependencies:

    [ontodev/excel "0.2.5"]

Then `require` the namespace:

    (ns your.project
      (:require [ontodev.excel :as xls]))

Use it to load a workbook and read sheets:

    (let [workbook (xls/load-workbook "test.xlsx")
          sheet    (xls/read-sheet workbook "Sheet1")]
      (println "Sheet1:" (count sheet) (first sheet)))

## License

Copyright © 2014, James A. Overton

Distributed under the Simplified BSD License: [http://opensource.org/licenses/BSD-2-Clause]()

