class MonthEndFile:
    def __init__(self, name): #, tabs=None, checks=0, details = {}, errors = []):
        self.name = name
        self.tabs = 0
        self.checks = 0
        self.details = {}
        self.errors = []

    def check_value(self, varname, varvalue):
        try:
            if self.details[varname] != varvalue:
                self.errors.append("Discrepancy in %s used between tabs" %varname)
                print("\t Discrepancy in %s used between tabs" %varname)
        except:
            self.details[varname] = varvalue

    
    def check_complete(self):
        self.checks += 1