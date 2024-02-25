#from openpyxl import load_workbook
#from openpyxl.utils import datetime
from datetime import datetime, timedelta

class KacPlanSheet:
    PLACE_SYMBOLS = ['SM', 'S', 'M', 'N']
    def __init__(self, conf, sheet, id):
        self.conf = conf
        self.sheet = sheet
        self.id = id
        self.date_value = sheet['D2'].value # D2:年月
        self.plan_list = []

    def print_plan_list(self):
        if self.conf["verbose"]:
            print("[print_plan_list]")
        for plan in self.plan_list:
            print(plan)

    def get_plan_list(self):
        return self.plan_list

    def get_date(self, day):
        if self.conf["verbose"]:
            print("[get_date]")
        day_delta = timedelta(days=day-1)
        combined_date = self.date_value + day_delta
        return combined_date.strftime("%Y-%m-%d")

    def get_symbol(self, places, target_place):
        if places[target_place] == 1:
            return self.PLACE_SYMBOLS[target_place]
        else:
            return None

    def make_plan(self, date, weekday, timesymbol, places ):
        if self.conf["verbose"]:
            print(f"[make_plan] timesymbol:{timesymbol}")
        for index in range(4):
            symbol = self.get_symbol(places, index)
            if symbol is not None:
                print(f"---", date, weekday, timesymbol, symbol, self.id)
                self.plan_list.append( [date, weekday, timesymbol, symbol, self.id])

    def get_position(self, ref_row, ref_column):
        if self.conf["verbose"]:
            print(f"[get_position] row:{ref_row} column:{ref_column}")
        data_list = []
        for column in range(ref_column, ref_column + 4):
            cell_value = self.sheet.cell(row=ref_row, column=column).value
            data_list.append(cell_value)
        return data_list

    def get_day_block(self, ref_row, ref_column):
        if self.conf["verbose"]:
            print(f"[get_day_block] row:{ref_row} column:{ref_column}")
        row=ref_row
        column = ref_column
        day_value = self.sheet.cell(row=ref_row, column=column).value
        weekday = self.sheet.cell(row=row+1, column=column).value
        if day_value is None:
            return
        #if day_value > 0 and day_value <= 31:
        ymd = self.get_date(day_value)
        #print(f"day: {ymd} weekday:{weekday}")
        row=ref_row+3
        places = self.get_position( row, column)
        self.make_plan(ymd, weekday, "T4", places)

        row=ref_row+4
        places = self.get_position( row, column)
        self.make_plan(ymd, weekday, "T5", places)

    def get_week_block(self, ref_row, ref_column):
        if self.conf["verbose"]:
            print(f"[get_week_block] row:{ref_row} column:{ref_column}")
        self.get_day_block(ref_row, ref_column)
        self.get_day_block(ref_row, ref_column+5)
        self.get_day_block(ref_row, ref_column+10)

    def get_month_block(self):
        if self.conf["verbose"]:
            print("[get_month_block]")
        self.get_week_block( 4, 4)
        self.get_week_block( 9, 4)
        self.get_week_block(14, 4)
        self.get_week_block(19, 4)
        self.get_week_block(24, 4)
