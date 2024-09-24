import requests
import pandas as pd

class TrialScrape:

    def fetch_data(self, nct_id, json_format=True):
        base_url = f'https://clinicaltrials.gov/api/v2/studies/{nct_id}'

        headers = {
            'Accept': 'application/json' if json_format else 'application/xml'
        }
        # Send the request to the API
        response = requests.get(base_url, headers=headers)
        if response.status_code == 200:
           
            if json_format:
                try:
                    self.data = response.json()

                except ValueError as e:
                    print(f"JSON decode error: {e}")
                    return None
            else:
                self.data = response.text  # XML response will be returned as text
            return self.data
        else:
            print(f"HTTP error occurred: {response.status_code}")
            return None

    def getOveralStatus(self):
        return self.data['protocolSection']['statusModule']['overallStatus']

    def getFirstPosted(self):
        return self.data['protocolSection']['statusModule']['studyFirstPostDateStruct']['date']

    def getName(self):
        return self.data['protocolSection']['identificationModule']['acronym']

    def getLastUpdate(self):
        return self.data['protocolSection']['statusModule']['lastUpdatePostDateStruct']['date']

    def getIndication(self):
        return self.data['protocolSection']['conditionsModule']['conditions']

    def getSSD(self):
        return self.data['protocolSection']['statusModule']['startDateStruct']['date']

    def getPCD(self):
        return self.data['protocolSection']['statusModule']['primaryCompletionDateStruct']['date']

    def getSCD(self):
        return self.data['protocolSection']['statusModule']['completionDateStruct']['date']

    def getPCDStatus(self):
        return self.data['protocolSection']['statusModule']['primaryCompletionDateStruct']['type']

    def getSponsor(self):
        return self.data['protocolSection']['sponsorCollaboratorsModule']['leadSponsor']['name']

    def getPhase(self):
        return self.data['protocolSection']['designModule']['phases']

    def getEnrolNo(self):
        return self.data['protocolSection']['designModule']['enrollmentInfo']['count']

    def getEnrolStatus(self):
        return self.data['protocolSection']['designModule']['enrollmentInfo']['type']

    def getPopulation(self):
        return self.data['protocolSection']['eligibilityModule']['stdAges']

    def getSetting(self):
        ages = []
        ages.append(self.data['protocolSection']['eligibilityModule']['minimumAge'])
        if 'maximumAge' in self.data['protocolSection']['eligibilityModule']:
            ages.append(self.data['protocolSection']['eligibilityModule']['maximumAge'])
        return ages

    def getCriteria(self):
        return self.data['protocolSection']['eligibilityModule']['eligibilityCriteria']

    def getPrimaryEndpoints(self):
        outcomes = []
        for endpoint in self.data['protocolSection']['outcomesModule']['primaryOutcomes']:
            outcomes.append(endpoint)
        return outcomes

    def getArms(self):
        arms = []
        for arm in self.data['protocolSection']['armsInterventionsModule']['armGroups']:
            arms.append(arm['label'] +  ":"  + arm['type'] +  ":"  + arm['description'])
        return arms


df = pd.read_excel(r"[SOURCE EXCEL SET UP IN THE FORMAT AS OUTLINED IN README].xlsx")


trials = df.to_dict()
print(trials)
ts = TrialScrape()
for i in range(len(trials["Registry Code"])):

    try:
        if trials["IncludeScrape"][i] == "Yes":

            nct_id = trials["Registry Code"][i]

            ts.fetch_data(nct_id)

            try:
                trials["First Posted"][i] = ts.getFirstPosted()

            except:
                pass

            try:
                trials["Name"][i] = ts.getName()
            
            except:
                pass

            try:
                trials["PCD"][i] = ts.getPCD()
            except:
                pass

            try:
                trials["Phase"][i] = ts.getPhase()

            except:

                pass

            try:
                trials["Last Updated"][i] = ts.getLastUpdate()
            except:
                pass

            try:
                trials["Status"][i] = ts.getOveralStatus()
            except:
                pass

            try:
                trials["Enrollment"][i] = ts.getEnrolNo()
            except:
                pass

            try:
                trials["SSD"][i] = ts.getSSD()
            except:
                pass

            try:
                trials["Enrollment Status"][i] = ts.getEnrolStatus()
            except:
                pass

            try:
                trials["SCD"][i] = ts.getSCD()
            except:
                pass

            try:
                trials["PCD Status"][i] = ts.getPCDStatus()
            except:
                pass

            try:
                trials["Arms"][i] = ts.getArms()
            except:
                pass

            try:
                trials["Primary Endpoint"][i] = ts.getPrimaryEndpoints()
            except:
                pass

            try:
                trials["Criteria"][i] = ts.getCriteria()
            except:
                pass

            try:
                    trials["Indication"][i] = ts.getIndication()
            except:
                continue

            try:
                trials['Setting'][i] = ts.getSetting()
            except:
                pass

            try:
                trials['Population'][i] = ts.getPopulation()
            except:
                pass

            try:
                trials["Sponsor"][i] = ts.getSponsor()
            except:
                continue

    except:
        continue



scrapedData = pd.DataFrame.from_dict(trials)

scrapedData.to_excel(r"[SOURCE EXCEL SET UP IN THE FORMAT AS OUTLINED IN README].xlsx", index = False)
