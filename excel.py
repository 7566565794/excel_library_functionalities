 def open_excel(self,args):
        wb = openpyxl.load_workbook(args[0])


    def create_excel_sheet(self,args):
        wb = openpyxl.load_workbook(args[0])
        wb.create_sheet(args[1])
        wb.save(args[0])

    def view_all_sheets(self,args):
        wb = openpyxl.load_workbook(args[0])
        print(wb.get_sheet_names())

    def switch_specific_sheets(self,args):
        wb = openpyxl.load_workbook(args[0])
        # s1=wb.active
        # print(s1.title)
        sheet = wb.get_sheet_by_name(args[1])
        print(sheet['A1'].value)


    def save_sheet(self,args):
        wb = openpyxl.load_workbook(args[0])
        wb.save(args[1])

    def desired_value(self,args):
        wb = openpyxl.load_workbook(args[0])
        sheet = wb.get_sheet_by_name(args[1])
        print(sheet.cell(row=int(args[2]), column=int(args[3])).value)

    def enter_value(self,args):
        wb = openpyxl.load_workbook(args[0])
        sheet = wb.get_sheet_by_name(args[1])
        sheet.cell(row=int(args[2]), column=int(args[3])).value = args[4]
        wb.save(args[0])

    def drop_sheet(self,args):
        wb = openpyxl.load_workbook(args[0])
        # sheet = wb.get_sheet_by_name(args[1])
        std = wb.get_sheet_by_name(args[1])
        # wb.save(args[0])
        wb.remove_sheet(std)
        wb.get_sheet_names()
        wb.save(args[0])

    # def get_active_sheet(self,args):
    #     wb = openpyxl.load_workbook(args)
    #     print(sheet.title)
    # def list_sheets(self,args):
