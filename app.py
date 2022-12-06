_author_ = 'arichland'

import Analyze
import query

filename = "Inventory Analysis.xlsx"

class InvenAnalysis:
    def __init__(self, **kargs):
        self.qry = query.query
        self.fc = Analyze.forecasts.forecasts
        self.cur = Analyze.inven_current.current
        self.hist = Analyze.inven_history.history
        self.analyze = [kargs['analyze']] if type(kargs['analyze']) == type(str()) else kargs['analyze']

    def arguments(self):
        if self.analyze == True:
            pass

        elif type(self.analyze) == type(list()):
            for analysis in self.analyze:
                if analysis == 'forecasts':
                    self.fc(calc=True)

                elif analysis == 'current':
                    self.cur(calc=True)

                elif analysis == 'inventory history':
                    self.hist(calc=True)

def main(**kargs):
    go = InvenAnalysis(**kargs)
    go.arguments()

if __name__ == '__main__':
    main(analyze='forecasts')