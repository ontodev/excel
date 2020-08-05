(defproject ontodev/excel "0.2.5"
  :description "A thin Clojure wrapper around a small part of Apache POI for
                reading .xlsx files."
  :url "http://github.com/ontodev/excel"
  :license {:name "Simplified BSD License"
            :url "http://opensource.org/licenses/BSD-2-Clause"}
  :dependencies [[org.clojure/clojure "1.10.1"]
                 [org.clojure/tools.logging "0.2.6"]
                 [org.apache.poi/poi-ooxml "4.1.2"]]
  :profiles
  {:dev {:dependencies [[midje "1.9.9"]
                        [lazytest "1.2.3"]]
         :plugins [[lein-midje "3.1.3"]
                   [lein-marginalia "0.7.1"]]}})
