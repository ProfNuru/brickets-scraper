from pathlib import Path
import datetime

import scrapy
import pandas as pd
from openpyxl import load_workbook

class BricksetSpider(scrapy.Spider):
    name="brickset"

    def start_requests(self):
        year = int(datetime.date.today().strftime("%Y"))

        urls = [f'https://brickset.com/sets/theme-Speed-Champions/year-{year-x}' for x in range(0,11)]

        # urls = [
        #     'https://brickset.com/sets/theme-Speed-Champions/year-2023',
        # ]

        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def from_excel(self,pth):
        try:
            excel_data = pd.read_excel(pth)
            return excel_data
        except:
            df = pd.DataFrame.from_dict([])
            df.to_excel(pth)
            self.from_excel(pth)
    
    def parse(self,response):
        articles = response.css('article.set')
        meta = articles.css('.meta')
        legos = []
        file_path = 'brickset_10_years.xlsx'
        excel_data = self.from_excel(file_path)
        print(excel_data)
        for details in meta:
            launch_exit = details.css(".col>dl>dt:contains('Launch/exit') + dd::text").get()
            rrp = details.css(".col>dl>dt:contains('RRP') + dd::text").get()
            legos.append({
                "name":details.css('h1>a::text').get(),
                "launch_exit":launch_exit,
                "rrp":rrp if rrp else ""
            })
        if excel_data.empty:
            df = pd.DataFrame.from_dict(legos)
            df.to_excel(file_path,index=False)
        else:
            df = pd.DataFrame.from_dict(legos)
            updated_excel = pd.concat([excel_data,df])
            updated_excel.to_excel(file_path,index=False)
            # sheet_name = 'Sheet1'
            # with pd.ExcelWriter(file_path,engine='openpyxl', mode='a') as writer:
            #     writer.book = load_workbook(file_path)
            #     writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
                
