
import datetime
from collections import ChainMap

from pptx.util import Inches



class TrialPlacement:
    def __init__(self, monthlist, inches, year_list, startpoint,start_year):
        self.inches = inches
        self.years = len(year_list)
        self.monthlist = monthlist
        self.startpoint = startpoint
        self.start_year = datetime.datetime(start_year, 1, 1).strftime("%Y")

    def unit_calc(self):
        m = 12 * self.years
        self.unit = self.inches / m
        return self.unit

    def month_pos(self):
        mo = []
        for m in self.monthlist:
            if m == self.monthlist[0]:
                self.startpoint = self.startpoint
            else:
                self.startpoint = self.startpoint + self.unit
            dic = {m: self.startpoint}
            mo.append(dic)
        self.months = dict(ChainMap(*mo))
        return self.months

    def placement(self, start, end,i):
        first_date = self.monthlist[0]
        last_date = self.monthlist[-1]

        print(start,end,i)
        if start < first_date and end > last_date:
            L = Inches(self.months[first_date])
            w = Inches(self.months[last_date] - self.months[first_date]) + Inches(self.unit)
        elif start < first_date and end < last_date and end > first_date and str(self.start_year) not in end:
            L = Inches(self.months[first_date])
            w = Inches(self.months[end]) + Inches(self.unit) - L
        elif not start < first_date and end > last_date:
            L = Inches(self.months[start])
            w = Inches(self.months[last_date] - (self.months[start])) + Inches(self.unit)
        elif start < first_date and end < first_date:
            L = Inches(self.months[first_date])
            w = Inches(self.months[self.monthlist[0]]) + Inches(self.unit) - L
        elif start < first_date and str(self.start_year) in end:
            L = Inches(self.months[first_date])
            w = Inches(self.months[end] - self.months[first_date]) + Inches(self.unit)
        else:
            L = Inches(self.months[start])
            w = Inches(self.months[end] - self.months[start]) + Inches(self.unit)
        return (L, w)




