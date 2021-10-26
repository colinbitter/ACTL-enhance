# ACTL-enhance

ACTL-enhance.py targets xlsx files in the downloads folder in windows. Expected inputs:<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1) authorityList.xlsx (ACTL exported from Alma)<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2) ACTLp.xlsx (Analytics--all physical records created/modified in the last month)<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3) ACTLe.xlsx (Analytics--all electronic records created/modified in the last month)<br/>
Analytics data can also be pulled down via APIs<br/><br/>
ACTL-enhance.py queries id.loc.gov to account for timing issues between LCNAF and the Alma CZ, eliminates headings which do not require authority control, and adds extra information to the ACTL for user convenience. Data is output to xlsx in the downloads folder.
