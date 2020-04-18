(defproject ontodev/excel "0.2.5"
  :description "A thin Clojure wrapper around a small part of Apache POI for
                reading .xlsx files."
  :url "http://github.com/ontodev/excel"
  :license {:name "Simplified BSD License"
            :url "http://opensource.org/licenses/BSD-2-Clause"}
  :dependencies [[org.clojure/clojure "1.5.1"]
                 [org.clojure/tools.logging "0.2.6"]
                 [org.apache.poi/poi-ooxml "3.8"]]
  :profiles
  {:dev {:dependencies [[midje "1.6.3"]
                        [lazytest "1.2.3"]]
         :plugins [[lein-midje "3.1.3"]
                   [lein-marginalia "0.7.1"]]}})
