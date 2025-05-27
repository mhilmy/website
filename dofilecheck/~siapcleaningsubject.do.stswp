clear
set more off
capture log close

global IslEd "D:/Dropbox/Islamic Education/Data/SIAP"
global IslEd "~/Dropbox/Islamic Education/Data/SIAP"

cd "$IslEd"
/*
insheet using "$IslEd/rawexcel/sekolah_subjects.csv", double clear

import excel using "$IslEd/rawexcel/sekolah_timetable_v2.xlsx", firstrow clear
recast strL Senin-Minggu
compress
saveold "$IslEd/sekolah_timetable_v2", replace
import excel using "$IslEd/rawexcel/sekolah_timetable.xlsx", firstrow clear
recast strL Senin-Minggu
compress
saveold "$IslEd/sekolah_timetable", replace
import excel using "$IslEd/rawexcel/sekolah_profile_v2.xlsx", firstrow clear
saveold "$IslEd/sekolah_profile_v2", replace

use "$IslEd/sekolah_timetable_v2",clear
append using "$IslEd/sekolah_timetable"
duplicates drop
missings dropobs Senin-Minggu
duplicates drop URL Tingkat Rombel, force
label data "Timetable data, combined v1 and v2 scrape"
saveold "$IslEd/sekolah_profile_combined", replace

** NEED TO ADD AGAIN MISSINGS DROPOBS AND APPEND DUPLICATE DROP
*/
use "$IslEd/sekolah_timetable_combined", clear
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
	
	* THIS LINE IS NEW
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


*use "$IslEd/sekolah_timetable_clean_v2", clear
*use "$IslEd/sekolah_timetable_clean_combined", clear
use "$IslEd/SIAP/sekolah_timetable_clean_combined", clear

preserve
* Get tagged list
//if in cluster it is "$IslEd/Subjectlist_tagged.xlsx"
import excel using "$IslEd/SIAP/Subjectlist_tagged.xlsx", clear firstrow
tempfile subjtag
save `subjtag'
import excel using "$IslEd/SIAP/SUBJECT_TAG_V2.xlsx", clear
rename (B C) (subjectname islamsubjecttype)
replace islamsubjecttype = "SEJARAH" if inlist(islamsubjecttype,"AHLI SUNAH WAL JAMA'AH", ///
	"AL KHAIRAAT","MUHAMMADIYAH","NAHDLATUL ULAMA","NAHDLATUL WATHAN","AL WASHLIYAH")
replace islamsubjecttype = "GENERAL ISLAM" if inlist(islamsubjecttype,"DAKWAH","LOGIKA","KALIGRAFI","KITAB KUNING","UNCATEGORIZED")
replace islamsubjecttype = "FIKIH" if islamsubjecttype == "DOA" 
tempfile subjtag2
save `subjtag2'
restore
*merge m:1 subjectname using `subjtag2', nogenerate keep(match master)
*rename senin1_subject subjectname
*merge m:1 subjectname using `subjtag', nogenerate keep(match master)
*merge m:1 subjectname using `subjtag2', nogenerate keep(match master)


foreach hari in senin selasa rabu kamis jumat sabtu minggu{
ds `hari'*_subject, v(32)
local `hari'_count : word count `r(varlist)'
display "``hari'_count'"

forval i = 1/``hari'_count' {
  rename `hari'`i'_subject subjectname
  * Tag subject as islamic/not
  merge m:1 subjectname using `subjtag', nogenerate keep(match master) //to get indicator of islamic subjects
  merge m:1 subjectname using `subjtag2', nogenerate keep(match master) //to get type of islamic subject
  rename (subjectname islamsubject islamsubjecttype) (`hari'`i'_subject `hari'`i'_islam `hari'`i'_islamtype)
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
  
  gen `hari'`i'_akidahduration = `hari'`i'_duration if `hari'`i'_islamtype == "AKIDAH AKHLAK"
  gen `hari'`i'_quranduration = `hari'`i'_duration if `hari'`i'_islamtype == "AL QURAN HADIS"
  gen `hari'`i'_bhsarabduration = `hari'`i'_duration if `hari'`i'_islamtype == "BAHASA ARAB"
  gen `hari'`i'_fikihduration = `hari'`i'_duration if `hari'`i'_islamtype == "FIKIH"
  gen `hari'`i'_genislamduration = `hari'`i'_duration if `hari'`i'_islamtype == "GENERAL ISLAM"
  gen `hari'`i'_historyislamduration = `hari'`i'_duration if `hari'`i'_islamtype == "SEJARAH"
  order `hari'`i'_akidahduration `hari'`i'_quranduration `hari'`i'_bhsarabduration ///
	`hari'`i'_fikihduration `hari'`i'_genislamduration `hari'`i'_historyislamduration, ///
	after(`hari'`i'_salafduration) 
}

}

drop if missing(senin1_subject) & mi(selasa1_subject) & mi(rabu1_subject) & ///
        mi(kamis1_subject) & mi(jumat1_subject) & mi(sabtu1_subject) & mi(minggu1_subject)

foreach hari in senin selasa rabu kamis jumat sabtu minggu{
  egen `hari'_totalduration = rowtotal(`hari'*_duration)
  foreach sub in islam ppkn pjok salaf akidah quran bhsarab fikih genislam historyislam{
	egen `hari'_total`sub'duration = rowtotal(`hari'*_`sub'duration)
	}
}

egen weekly_totalduration = rowtotal(*totalduration)
foreach sub in islam ppkn pjok salaf akidah quran bhsarab fikih genislam historyislam{
	egen weekly_total`sub'duration = rowtotal(*total`sub'duration)
	gen weekly_`sub'share = weekly_total`sub'duration / weekly_totalduration
}

order weekly_* *_totalduration *_total*duration, before(senin1_subject)
compress
missings dropvars, force
recast strL *_subject *_islamtype
compress
saveold sekolah_timetable_clean_combined_tagged, replace

use sekolah_profile_v2, clear
append using sekolah_profile
missings dropobs, force
duplicates drop
egen missingfield = rowmiss(_all)
gsort URL missingfield 
by URL: gen n = _n
drop if n != 1
drop n missingfield
recast strL guru Nama Akreditasi Alamat NomorTelpon NomorFaks Email Situs Lintang Bujur WaktuBelajar Kota-Kelurahan
compress
save sekolah_profile_combined, replace

use sekolah_profile_combined, clear
rename URL website
merge 1:1 website using SIAP, keep(match master) nogenerate keepusing(madrasah islam)
rename website url
merge 1:m url using sekolah_timetable_clean_combined_tagged

 gen totalhours = totalduration/60
 format %4.0f totalhours
  rename weekly_* *
  format %5.2f *share
  
tabstat totalhours *share if madrasah == 1, by(tingkat) format
tabstat totalhours *share if islam == 1 & madrasah ==0, by(tingkat) format
tabstat totalhours *share if inlist(1,islam,madrasah), by(tingkat) format

label data "SIAP scrape timetable merged with school profile, v1 and v2 scrapes combined"
save SIAP_profile_timetable_tagged, replace

