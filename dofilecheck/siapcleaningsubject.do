clear
set more off
capture log close

global IslEd "D:/Dropbox/Islamic Education/Data"
global IslEd "~/Dropbox/Islamic Education/Data"

*MH note Mar2024: This code updates 01i_clean_SIAPsubj_v2-sam.do (NB below) to add sci, math, etc.

*NB: this is Sam's code from December. Seems Masyhur updated the code later to calculate some of the subject-specific measures.
*	I believe they are the same as mine but just combine the different batches whereas I had them separate before...


cd "$IslEd"
/*
insheet using "$IslEd/rawexcel/sekolah_subjects.csv", double clear

import excel using "$IslEd/rawexcel/sekolah_timetable_v2.xlsx", firstrow clear
saveold "$IslEd/sekolah_timetable_v2", replace
import excel using "$IslEd/rawexcel/sekolah_profile_v2.xlsx", firstrow clear
saveold "$IslEd/sekolah_profile_v2", replace
*/
use "$IslEd/SIAP/sekolah_timetable_v2", clear
*use "/Users/masyhur/Dropbox/Islamic Education/Data/SIAP/sekolah_timetable1pct.dta", clear
rename *, lower
destring tingkat, replace
foreach hari in senin selasa rabu kamis jumat sabtu minggu{
  replace `hari' = upper(`hari')  
  replace `hari' = subinstr(`hari',"PENDIDIKAN JASMANI, OLAHRAGA, DAN KESEHATAN","PJOK",.)
  replace `hari' = subinstr(`hari',"TEMATIK PJOK","PJOK",.)
  replace `hari' = "PJOK" if regexm(`hari',"OLAHRAGA")==1
  replace `hari' = subinstr(`hari',"PENDIDIKAN PANCASILA DAN KEWARGANEGARAAN","PPKN",.)
  replace `hari' = subinstr(`hari',",,",",",.) 
  replace `hari' = subinstr(`hari',", ,A"," A",.) //a case of "Fulan, ,A.Md"
  *foreach let in S A I D P M Q K F L N G H R Z B T .{ // to catch e.g. "Fulan, S.Pd, M.Pd, LC"
  foreach let in A B C D E F G H I J K L M N O P Q R S T U V W X Y Z .{ // to catch e.g. "Fulan, S.Pd, LC"
    replace `hari' = subinstr(`hari',", `let'"," `let'",.) //space between comma and gelar
    replace `hari' = subinstr(`hari',",`let'" ," `let'",.) //no space b/w comma and gelar
  }
  
  split `hari', parse(",")
  local `hari'_jum `r(nvars)'
  display "`hari'_jum ``hari'_jum'"
  
  rename (`hari'??) (`hari'??_) //this allows split to assign a new variable name
  rename (`hari'?) (`hari'?_)
  forval i=1/``hari'_jum'{
	display "hari `hari' period `i'"
	split `hari'`i'_, parse("|") //senin1_1 (hour+subj) dan senin1_2 (teacher)
	*display "Drop teacher name"
	*drop `hari'`i'_2
	*display "drop `hari'`i'_2"
	replace `hari'`i'_1 = subinstr(`hari'`i'_1,"(","",.) //remove open ( in start hour
	replace `hari'`i'_1 = trim(itrim(`hari'`i'_1))
	display "Get name of subject taught in period `i' on day `hari'"
	split `hari'`i'_1, parse(")") //senin1_12 (hour)
	replace `hari'`i'_12 = trim(itrim(`hari'`i'_12))
	rename  `hari'`i'_12 `hari'`i'_subject
	
	display "Calculate length of period `i' on day `hari'"
	split `hari'`i'_11, parse(" - ") // split period to start and end 
	gen `hari'`i'_111c = clock(`hari'`i'_111,"hm") //convert start period to stata time format
	gen `hari'`i'_112c = clock(`hari'`i'_112,"hm") //convert end period to stata time format
	format %tc `hari'`i'_111c `hari'`i'_112c
	gen `hari'`i'_duration =  minutes(`hari'`i'_112c-`hari'`i'_111c)
	
	replace `hari'`i'_duration = minutes(tc(2jan1960 00:00)-`hari'`i'_111c) if `hari'`i'_duration < 0
	drop `hari'`i'_112* `hari'`i'_111*
  } //END BRACE FORVAL 1-END OF PERIOD IN A DAY
} //ENDBRACE FOR EACH DAY

keep url-rombel *subject *duration
compress
saveold sekolah_timetable_clean_v2, replace


* Generate subject lists for manual categorizing whether subject is islamic/not
keep *subject
duplicates drop
gen i = _n
rename *_subject *
reshape long senin selasa rabu kamis jumat sabtu minggu,i(i) j(period)
rename (senin selasa rabu kamis jumat sabtu minggu) (hari1 hari2 hari3 hari4 hari5 hari6 hari7)
drop i period
gen i = _n
drop if missing(hari1) &  missing(hari2) &  missing(hari3) &  missing(hari4) &  ///
	missing(hari5) &  missing(hari6) &  missing(hari7)
reshape long hari, i(i) j(day)
drop i day
replace hari = upper(hari)
sort hari
duplicates drop
drop if missing(hari)
rename hari subject
export excel using "Subjectlist_v2.xlsx", replace

*--->NOTE: we need to clean the above if subject list changes in batch 2 

* Get tagged list
import excel using "$IslEd/SIAP/Subjectlist_tagged.xlsx", clear firstrow
tempfile subjtag
save `subjtag'


use SIAP/sekolah_timetable_clean_v2,clear
foreach hari in senin selasa rabu kamis jumat sabtu minggu{
ds `hari'*_subject, v(32)
local `hari'_count : word count `r(varlist)'
display "``hari'_count'"

forval i = 1/``hari'_count' {
  rename `hari'`i'_subject subjectname
  * Tag subject as islamic/not
  merge m:1 subjectname using `subjtag', nogenerate keep(match master)
  rename (subjectname islamsubject) (`hari'`i'_subject `hari'`i'_islam)
  * Generate duration variable specific to islam subjects only
  gen `hari'`i'_islamduration = `hari'`i'_duration if `hari'`i'_islam == 1
  order `hari'`i'_islam `hari'`i'_islamduration, after(`hari'`i'_duration)
  * Generate duration variable specific to PPKN
  gen `hari'`i'_ppkn = `hari'`i'_subject == "PPKN" if `hari'`i'_subject!=""
  gen `hari'`i'_ppknduration = `hari'`i'_duration if `hari'`i'_subject == "PPKN"
  order `hari'`i'_ppkn `hari'`i'_ppknduration, after(`hari'`i'_islamduration)
  * Generate duration variable specific to PJOK
  gen `hari'`i'_pjok = `hari'`i'_subject == "PJOK" if `hari'`i'_subject!=""
  gen `hari'`i'_pjokduration = `hari'`i'_duration if `hari'`i'_subject == "PJOK"
  order `hari'`i'_pjok `hari'`i'_pjokduration, after(`hari'`i'_ppknduration)
  * Generate duration variable specific to Salafi subjects
  gen `hari'`i'_salaf = regexm(`hari'`i'_subject,"SALAF") if `hari'`i'_subject!=""
  gen `hari'`i'_salafduration = `hari'`i'_duration if regexm(`hari'`i'_subject,"SALAF")
  order `hari'`i'_salaf `hari'`i'_salafduration, after(`hari'`i'_pjokduration)
  * Generate duration variable specific to Bahasa Indo subjects
  gen `hari'`i'_indo = regexm(`hari'`i'_subject,"INDO") & regexm(`hari'`i'_subject,"BAHASA") if `hari'`i'_subject!=""
  gen `hari'`i'_indoduration = `hari'`i'_duration if regexm(`hari'`i'_subject,"INDO") & regexm(`hari'`i'_subject,"BAHASA")
  order `hari'`i'_indo `hari'`i'_indoduration, after(`hari'`i'_salafduration)
  * Generate duration variable specific to Bahasa Arab subjects
  gen `hari'`i'_arab = regexm(`hari'`i'_subject,"ARAB") & !regexm(`hari'`i'_subject,"FIQ") & !regexm(`hari'`i'_subject,"FIK")  if `hari'`i'_subject!=""
  gen `hari'`i'_arabduration = `hari'`i'_duration if regexm(`hari'`i'_subject,"ARAB") & !regexm(`hari'`i'_subject,"FIQ") & !regexm(`hari'`i'_subject,"FIK") 
  order `hari'`i'_arab `hari'`i'_arabduration, after(`hari'`i'_indoduration)  
  * Generate duration variable specific to Math subjects
  gen `hari'`i'_math = regexm(`hari'`i'_subject,"MATEMATIKA")  if `hari'`i'_subject!=""
  gen `hari'`i'_mathduration = `hari'`i'_duration if regexm(`hari'`i'_subject,"MATEMATIKA") 
  order `hari'`i'_math `hari'`i'_mathduration, after(`hari'`i'_arabduration)
  * Generate duration variable specific to Sci subjects
  gen `hari'`i'_sci = regexm(`hari'`i'_subject,"ILMU PENGETAHUAN ALAM") | ///
			regexm(`hari'`i'_subject,"FISIKA")| regexm(`hari'`i'_subject,"KIMIA")| ///
			regexm(`hari'`i'_subject,"BIOLOGI") if `hari'`i'_subject!=""
  gen `hari'`i'_sciduration = `hari'`i'_duration if `hari'`i'_sci == 1
  order `hari'`i'_sci `hari'`i'_sciduration, after(`hari'`i'_mathduration)
  * Generate duration variable specific to humanities subjects
  gen `hari'`i'_hum = regexm(`hari'`i'_subject,"ILMU PENGETAHUAN SOSIAL") | ///
			regexm(`hari'`i'_subject,"EKONOMI")| regexm(`hari'`i'_subject,"SEJARAH")| ///
			regexm(`hari'`i'_subject,"SOSIOLOGI")| regexm(`hari'`i'_subject,"WIRAUSAHA")| ///
			regexm(`hari'`i'_subject,"ANTROPOLOGI")| regexm(`hari'`i'_subject,"GEOGRAFI") if `hari'`i'_subject!=""
  gen `hari'`i'_humduration = `hari'`i'_duration if `hari'`i'_hum == 1
  order `hari'`i'_hum `hari'`i'_humduration, after(`hari'`i'_sciduration)
  * Generate duration variable specific to tematik subjects
  gen `hari'`i'_tema = regexm(`hari'`i'_subject,"^TEMATIK ")  if `hari'`i'_subject!=""
  gen `hari'`i'_temaduration = `hari'`i'_duration if `hari'`i'_tema == 1
  order `hari'`i'_tema `hari'`i'_temaduration, after(`hari'`i'_humduration)
}

}

drop if missing(senin1_subject) & mi(selasa1_subject) & mi(rabu1_subject) & ///
        mi(kamis1_subject) & mi(jumat1_subject) & mi(sabtu1_subject) & mi(minggu1_subject)

foreach hari in senin selasa rabu kamis jumat sabtu minggu{
  egen `hari'_totalduration = rowtotal(`hari'*_duration)
  foreach sub in islam ppkn pjok salaf indo arab math sci hum tema{
	egen `hari'_total`sub'duration = rowtotal(`hari'*_`sub'duration)
	}
}

egen weekly_totalduration = rowtotal(*totalduration)
foreach sub in islam ppkn pjok salaf indo arab math sci hum tema{
	egen weekly_total`sub'duration = rowtotal(*total`sub'duration)
	gen weekly_`sub'share = weekly_total`sub'duration / weekly_totalduration
}

order weekly_* *_totalduration *_total*duration, before(senin1_subject)
quietly compress
missings dropvars, force
keep url tahunajaran semester tingkat rombel weekly*

saveold "$IslEd/SIAP/sekolah_timetable_clean_v3_tagged", replace
use "$IslEd/SIAP/sekolah_timetable_clean_v3_tagged", clear

**//calculate share of unstructured subject names not already captured. Most are captured here.

use *_subject using "$IslEd/sekolah_timetable_clean_v2", clear 
gen i = _n
rename *_subject *
reshape long senin selasa rabu kamis jumat sabtu minggu,i(i) j(period)
rename (senin selasa rabu kamis jumat sabtu minggu) (hari1 hari2 hari3 hari4 hari5 hari6 hari7)
drop i period
gen i = _n
drop if missing(hari1) &  missing(hari2) &  missing(hari3) &  missing(hari4) &  ///
	missing(hari5) &  missing(hari6) &  missing(hari7)
reshape long hari, i(i) j(day)
drop i day
replace hari = upper(hari)
sort hari
drop if missing(hari)
rename hari subjectname
gen n = 1
collapse (sum) n, by(subjectname)
gsort -n
gen sum = n if _n == 1
gen sum2 = n if _n == 2

replace sum2 = sum2[_n-1]+n if mi(sum2)
replace sum2 = sum2/5905267
replace sum = sum[_n-1]+n if mi(sum)
gen share = sum/5905267


preserve
import excel using "$IslEd/Subjectlist_tagged.xlsx", clear firstrow
tempfile subjtag
save `subjtag'
restore
merge m:1 subjectname using `subjtag', nogenerate keep(match master)
