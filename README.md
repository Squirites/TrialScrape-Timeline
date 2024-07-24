The code provided here scrapes information from ClinicalTrials.gov using the ClinicalTrials.gov API into an Excel (executed in trialscrape.py).

A trial with an NCT-code can be scraped, however trials without NCT codes cannot be scraped. For a trial with an NCT code, ensure that 'Yes' is in the respective row of the 'IncludeScrape' column. For trials without an NCT number, information can be included manually, 
but you need to ensure that the following columns are not null (below). For manually inputted trials, leave the 'IncludeScrape' column blank.

PPT slides presenting timelines and other trial information from that Excel can then be generated using the other code (executed in main). Further error handling is yet to be added to the code. To ensure that no errors arise during the generation of the timelines,
ensure that the following columns are not empty / null for each trial:

Registry Code;
Name;
Therapy;
Sponsor;
RoA;
Mechanism of Action;
Population;
Setting;
Indication;
Phase;
Status;
Enrollment;
Link;
SSD;
PCD;

The remaining columns can be left blank and it will not impact the generation of the slides

If you are using a specific template you may need to adjust the dimensions in the Main.py file to ensure the format is correct. You can also change the time period on which the trials are plotted, by changing the start_year and end_year variables in Main.py - currently they are set to 2021 and 2027 respectively.
