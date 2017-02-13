+++++++++++++++++++++++++++++++++++++++++++++++++++

OTD:
1. copy the input files into Data-->OTD folder
2. change the file names per below
Input1-TL9000
Input2-TW-Close
Input3-TW-Open
Input4-KOREA
Input5-NZ
3. delete the non-data header line & bottom line
+++++++++++++++++++++++++++++++++++++++++++++++++++


NonDesp:
1. copy the input file into Data-->NonDesp
2. change the file name to Input6-NonDesp
3. delete the non-data header line & bottom line
+++++++++++++++++++++++++++++++++++++++++++++++++++


Monthly Open Order Report:
1. copy the input files into Data-->OpenOrder
2. change the file name per below
Input7-OpenOrderOnelog
Input8-NZ (same file as OTD NZ. Convertor auto-chooses a different worksheet in it to use.)
+++++++++++++++++++++++++++++++++++++++++++++++++++


Weekly OTD key customer performance:
1. copy the done file of weekly OTD into Data-->WeeklyOTDAnalysis
2. change the file name to WeeklyOTD
3. delete the non-data header lines
+++++++++++++++++++++++++++++++++++++++++++++++++++


Monthly OTD top 10 customer & TW customer name translation:
1. copy the done file of monthly OTD (4 weeks' accumulation) into Data-->MonthlyOTDAnalysis
2. change the file name to MonthlyOTD
3. delete the non-data header lines
+++++++++++++++++++++++++++++++++++++++++++++++++++


Monthly closed RMA report:
1. copy the 5 VC files into Data-->closedRMA-->Ref
2. change the file names as below:
VC Catalogue - RESO AMERICAS NAR
VC Catalogue - RESO APAC CHINA QD
VC Catalogue - RESO APAC INDIA
VC Catalogue - RESO China (SHA)
VC Catalogue - RESO EMEA

3. In RLC-India VC file, remove the stop sign in column name 'RES_ID'

4. copy the 5 source files into Data-->ClosedRMA
5. change the file names as below:
ClosedRMA-ASB
ClosedRMA-Citadel Report
ClosedRMA-NZ
ClosedRMA-Ormes
ClosedRMA-eSpares TW

6. change the active worksheet into 'Sheet1' for Citadel & RGL-NZ
7. delete the non-data header line for TW
8. In RGL-NZ file, align the relevant columns' names:
column AR --> 'repairer1' 
column AS --> 'RMA1'
column Ax --> 'repairer2' 


issues:
1) QD VC file, the sheet name has changed from 'RES China (QD) VC' to 'ASIA PACIFIC' since 3Q2011
2) QD VC file, the column names changes also: 'VC_RFR60' --> ITP_RFR_60, 'VC_NFF' -->ITP_NFF
3) Ormes VC, column 'VC_NFF' changed to 'VC_RFR_NFF' since 3Q2011
4) in Citadel file, 3745 should be searched as cisco3745 in NAR VC
5) if no result in ASB VC, QD VC should be re-searched before assigning 215
6) India VC doesn't work to locate 3BK27162AA,3BK27175AA (Citadel file)
