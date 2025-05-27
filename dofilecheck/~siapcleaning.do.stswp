clear
set more off
global IDdata "/Users/masyhur/Dropbox/ID-data"
global IslEd "/users/masyhur/dropbox/islamic education/data"	

/*import excel using ~/Desktop/sekolah_Java.xlsx, firstrow
replace Jenjang = "SD" if Jenjang == "Sekolah Dasar"
replace Jenjang = "SMP" if Jenjang == "Sekolah Menengah Pertama"
replace Jenjang = "SMA" if Jenjang == "Sekolah Menengah Atas"
replace Jenjang = "SMK" if Jenjang == "Sekolah Menengah Kejuruan"
replace Jenjang = "TK" if Jenjang == "Taman Kanak Kanak"
replace Jenjang = "SLB" if Jenjang == "Sekolah Luar Biasa"
replace Jenjang = "ST" if Jenjang == "Sekolah Tinggi"*/

import excel using "$IslEd/SIAP/sekolah_all.xlsx", firstrow clear
replace Jenjang = "SD" if Jenjang == "Sekolah Dasar"
replace Jenjang = "SMP" if Jenjang == "Sekolah Menengah Pertama"
replace Jenjang = "SMA" if Jenjang == "Sekolah Menengah Atas"
replace Jenjang = "SMK" if Jenjang == "Sekolah Menengah Kejuruan"
replace Jenjang = "TK" if Jenjang == "Taman Kanak Kanak"
replace Jenjang = "SLB" if Jenjang == "Sekolah Luar Biasa"
replace Jenjang = "ST" if Jenjang == "Sekolah Tinggi"
drop if regexm(SchoolName,"Demo")
drop Logo

replace SchoolName = trim(itrim(upper(SchoolName)))
drop if regexm(SchoolName,"DEMO")
gen madrasah = .
replace madrasah = 1 if regexm(SchoolName,"^MI") & Jenjang == "SD"
replace madrasah = 1 if regexm(SchoolName,"^MTS") & Jenjang == "SMP"
replace madrasah = 1 if regexm(SchoolName,"^MA") & Jenjang == "SMA"
replace madrasah = 1 if regexm(SchoolName,"^MA") & Jenjang == "SMK"
replace madrasah = 1 if regexm(SchoolName,"^RA") & Jenjang == "TK"
replace madrasah = 1 if regexm(SchoolName,"^BA") & Jenjang == "TK"
replace madrasah = 1 if regexm(SchoolName,"MADRASAH")
replace madrasah = 0 if mi(madrasah)

gen islam = .
replace islam = 1 if madrasah == 1
replace islam = 1 if regexm(SchoolName,"ISLAM")
replace islam = 1 if regexm(SchoolName,"MUSLIM") 
foreach w in TK SD SMP SMA SMK{
	replace islam = 1 if regexm(SchoolName,"`w' IT")
	replace islam = 1 if regexm(SchoolName,"`w'IT")
}
replace islam = 1 if regexm(SchoolName,"SDI") & Status == "Swasta"
replace islam = 1 if regexm(SchoolName,"AZHAR")
replace islam = 1 if regexm(SchoolName,"MUHAMMADIYAH")
replace islam = 1 if regexm(SchoolName,"NURUL")
replace islam = 1 if regexm(SchoolName," NUR ")
replace islam = 1 if regexm(SchoolName,"SDS IT")
replace islam = 1 if regexm(SchoolName,"SDSI")
replace islam = 1 if regexm(SchoolName," AL-") & Status == "Swasta"
replace islam = 1 if regexm(SchoolName," AL ") & Status == "Swasta"
replace islam = 1 if regexm(SchoolName," AN-") & Status == "Swasta"
replace islam = 1 if regexm(SchoolName," AN ") & Status == "Swasta"
replace islam = 1 if regexm(SchoolName," AT-") & Status == "Swasta"
replace islam = 1 if regexm(SchoolName," AT ") & Status == "Swasta"
replace islam = 0 if mi(islam)

rename *, lower
gen siap_num = regexs(1) if regexm(website,"([0-9]+)")
encode prop, generate(propnum)
recode propnum (1=51) (2=36) (6=75)  ///
		(11=61) (12=63) (13=62) (14=64) (18=81) (19=82) ///15 16 17 
		(21=52) (22=53) (23=94) (24=91) (26=76) (27=73) (28=72) (29=74) (30=71) ///25
		(31=13) (33=12) (4=34)  (5=31)  (15=19) (16=21) (17=18) (25=14) (32=16) ///
		(3=17)  (7=15)  (8=32)  (9=33)  (10=35) (20=11)
label drop propnum
recode propnum (11/21=1) (31/36=2) (51/53=3) (61/64=4) (71/76=5) (81/94=6), generate(islands)
label define islandslab 1 "1-Sumatera" 2 "2-Java" 3 "3-NT" 4 "4-Kalimantan" ///
			5 "5-Sulawesi" 6 "6-Maluku/Papua"
label values islands islandslab		
label define madrasahlab 1 "madrasah" 0 "general"
label define islamlab 1 "islam" 0 "general"
label values madrasah madrasahlab
label values islam islamlab
label var islands ""
drop if  inlist(jenjang,"TK","SLB","ST")
gen jenjang2 = ""
replace jenjang2 = "1-SD" if jenjang == "SD"
replace jenjang2 = "2-SMP" if jenjang == "SMP"
replace jenjang2 = "3-SMA" if jenjang == "SMA"
replace jenjang2 = "4-SMK" if jenjang == "SMK"
encode jenjang2, generate(jenjangnum)
** generate statistics
tabulate islands status
table islands madrasah jenjang, row scolumn
table jenjang  islam madrasah, row scolumn

replace prop = upper(prop)
replace kotakab = upper(trim(itrim(kotakab)))
rename (prop kotakab) (prov kab)
drop if siap_num == ""
destring siap_num, gen(siap_num2)
save "$IslEd/SIAP/SIAP", replace


use  "$IslEd/madrasah/madrasah", clear
gen schoolname = itrim(trim(upper(jenjang+substr(status,1,1)+" "+nama)))
gen jenjang2=jenjang
replace jenjang = "SD" if jenjang == "MI"
replace jenjang = "SMP" if jenjang == "MTs"
replace jenjang = "SMA" if jenjang == "MA"
drop _merge
destring no, replace
tempfile madrasahsiap
save `madrasahsiap'

//beware this takes a really long time.
*reclink prov kab jenjang schoolname using "$IslEd/SIAP/SIAP", idm(no) idu(siap_num2) gen(siapmatch)

* Create two batches for further scrape work
use  "$IslEd/SIAP/SIAP", clear
set seed 23101012
keep if islam == 0
bysort prop jenjang status: sample 2, count
drop if inlist(jenjang,"TK","SLB","ST")
tempfile nonislamsch
save `nonislamsch'

use  "$IslEd/SIAP/SIAP", clear
keep if islam == 1
keep if jenjang != "TK"
drop if jenjang == "SLB"
append using `nonislamsch'
export excel schoolname prov kab jenjang status website using "$IslEd/SIAP/IslamSchoolList.xlsx", replace

gen rand = runiform()
gen batch1 = rand >0.5
export excel schoolname prov kab jenjang status website using "$IslEd/SIAP/IslamSchoolList1.xlsx" if batch1 == 1, replace
export excel schoolname prov kab jenjang status website using "$IslEd/SIAP/IslamSchoolList2.xlsx" if batch1 == 0, replace
*


* Load scraped dataset
import excel using ~/Desktop/sekolah_timetable.xlsx, clear firstrow

* NEW DATASETS NOV 25 2019

import excel using ~/Desktop/sekolah_timetable_v2.xlsx, clear firstrow
import excel using ~/Desktop/sekolah_profile_v2.xlsx, clear firstrow
import delimited using ~/Desktop/sekolah_students.csv, clear
import delimited using ~/Desktop/sekolah_subjects.csv, clear

* CREATE list of non-islamic schools for new scrape Jan 29 2020
use  "$IslEd/SIAP/SIAP", clear
keep if islam == 0
drop if inlist(jenjang,"TK","SLB","ST")
export excel schoolname prov kab jenjang status website using "$IslEd/SIAP/Non-RelgSchoolList.xlsx", replace
d, short
recast strL description
d, short
