The code provided here scrapes information from ClinicalTrials.gov using the ClinicalTrials.gov API into an Excel (executed in trialscrape.py), and then there is code to plot these trials into PowerPoint.

Instructions:

1. Download the template Excel, and population Registry Code column with relevant NCT trials for scraping. Only trials with an NCT-code can be automatically scraped. For a trial with an NCT code, ensure that 'Yes' is in the respective row of the 'IncludeScrape' column. For trials without an NCT number, information can be included manually, but you need to ensure that the following columns are not empty / null (below). For manually inputted trials, leave the 'IncludeScrape' column blank.

2. Scrape trial information using trialscrape.py in Trial Scrape folder

3. Ensure that for each row the following columns are not empty / null:

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

4. Plot trials using Main.py in Plot Trials Folder

      If you are using a specific template you may need to adjust the dimensions in the Main.py file to ensure the format is correct. You can also change the time period on which the trials are plotted, by changing the start_year and end_year variables in Main.py - currently   they are set to 2021 and 2027 respectively.
