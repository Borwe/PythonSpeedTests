import speedtest
import xlsxwriter
import os.path
import datetime

def check_or_createFile():
    # check if file exists, if not the create it
    file=datetime.datetime.now().strftime("%B")
    file+='-'+datetime.datetime.now().strftime("%Y")+'-speedtests.xlsx'

    if os.path.isfile(file) == True:
        print("File ",file," exists")
    else:
        print("File doesn't exists, creating it")
        workbook=xlsxwriter.Workbook(file)

        #creating worksheets
        lqd_worksheet=workbook.add_worksheet('LQD')
        jtl_worksheet=workbook.add_worksheet('JTL')
        saf_worksheet=workbook.add_worksheet('SAF')

        
        __fill_up_excel__(workbook,jtl_worksheet,'JTL')
        __fill_up_excel__(workbook,lqd_worksheet,'LQD')
        __fill_up_excel__(workbook,saf_worksheet,'SAF')
        
        workbook.close()

    return file

#fill up file with necessary titles for collumns
def __fill_up_excel__(workbook,worksheet,link_type):
    #morning text format
    cellB1=workbook.add_format()
    cellB1.set_italic()
    cellB1.set_underline()
    cellB1.set_bold()
    cellB1.set_bg_color('yellow')

    #BOLD big title fonts
    bold_bigTitle_format=workbook.add_format()
    bold_bigTitle_format.set_bold()
    bold_bigTitle_format.set_size(14)

    #add morning and afternoon text
    worksheet.write('B1','Morning',cellB1)
    worksheet.write('O1','Afternoon',cellB1)

    #create borders and titles
    bold_bigTitle_format.set_align('center')
    worksheet.merge_range('E2:H2',str(link_type+' LINK SPEEDTEST'),bold_bigTitle_format)
    worksheet.merge_range('R2:U2',str(link_type+' LINK SPEEDTEST'),bold_bigTitle_format)

    #Dark border format
    dark_border_format=workbook.add_format()
    dark_border_format.set_bold()
    dark_border_format.set_size(14)
    dark_border_format.set_border(2)
    dark_border_format.set_align('center')

    #BOLD TITLE WITH NORMAL BORDER FORMAT
    bold_title_normal_border=workbook.add_format()
    bold_title_normal_border.set_bold()
    bold_title_normal_border.set_size(12)
    bold_title_normal_border.set_border(1)
    bold_title_normal_border.set_align('center')

    #morning table
    worksheet.merge_range('C3:D3','UK',dark_border_format)
    worksheet.merge_range('E3:F3','US',dark_border_format)
    worksheet.merge_range('G3:H3','EUROPE',dark_border_format)
    worksheet.merge_range('I3:J3','NAIROBI',dark_border_format)
    #DATE raw titles
    worksheet.write('B4','DATE',bold_title_normal_border)
    worksheet.write('C4','Download',bold_title_normal_border)
    worksheet.write('D4','Upload',bold_title_normal_border)
    worksheet.write('E4','Download',bold_title_normal_border)
    worksheet.write('F4','Upload',bold_title_normal_border)
    worksheet.write('G4','Download',bold_title_normal_border)
    worksheet.write('H4','Upload',bold_title_normal_border)
    worksheet.write('I4','Download',bold_title_normal_border)
    worksheet.write('J4','Upload',bold_title_normal_border)
    worksheet.write('K4','Remarks',bold_title_normal_border)
    worksheet.write('L4','By',bold_title_normal_border)
    
    #evening table
    worksheet.merge_range('P3:Q3','UK',dark_border_format)
    worksheet.merge_range('R3:S3','US',dark_border_format)
    worksheet.merge_range('T3:U3','EUROPE',dark_border_format)
    worksheet.merge_range('V3:W3','NAIROBI',dark_border_format)
    #DATE raw titles
    worksheet.write('O4','DATE',bold_title_normal_border)
    worksheet.write('P4','Download',bold_title_normal_border)
    worksheet.write('Q4','Upload',bold_title_normal_border)
    worksheet.write('R4','Download',bold_title_normal_border)
    worksheet.write('S4','Upload',bold_title_normal_border)
    worksheet.write('T4','Download',bold_title_normal_border)
    worksheet.write('U4','Upload',bold_title_normal_border)
    worksheet.write('V4','Download',bold_title_normal_border)
    worksheet.write('W4','Upload',bold_title_normal_border)
    worksheet.write('X4','Remarks',bold_title_normal_border)
    worksheet.write('Y4','By',bold_title_normal_border)

class SerenaSpeedTester:

    def __init__(self):
        #hold Kenyan servers in list
        self.kenyan_servers=[]

        #hold United Kingdom servers in list
        self.uk_servers=[]
        
        #hold United States servers in list
        self.usa_servers=[]
        
        #Russia servers in list
        self.russia_servers=[]

        #speedtest object
        self.s=speedtest.Speedtest()

        self.__get_servers_based_on_our_four_region__()

        
    def __get_servers_based_on_our_four_region__(self):
        #get all speedtest.net servers
        self.__servers=self.s.get_servers()
        
        #display the ones in Kenya
        for point in self.__servers:
            self.server=self.__servers.get(point)
            #get country
            for part in self.server:
                # get servers that are in kenya
                if part.get('country').find('Kenya')!=-1:
                    self.kenyan_servers.append(part)
                     
                #get servers that are in UnitedKingdom
                if part.get('country').find('United Kingdom')!=-1:
                    self.uk_servers.append(part)

                #get servers that are in United States
                if part.get('country').find('United States')!=-1:
                    self.usa_servers.append(part)

                #get servers that are in United States
                if part.get('country').find('Russian Federation')!=-1:
                    self.russia_servers.append(part)

    #get list of IDs from server for usage
    def __get_country_servers_by_id__(self,country_servers):
        servers_by_id=[]
        for server in country_servers:
            servers_by_id.append(server.get('id'))

        return servers_by_id

    #turn bytes into Megabytes
    def __bytes_to_megabytes__(self,bytes):
        return round((bytes/(10**6)),2)

    #set time of checkup to evening
    def setTimeEvening(self):
        self.time='evening'

    #set time of checkup to morning
    def setTimeMorning(self):
        self.time='morning'

    #put data into file based on location in correct collumn
    #on correct worksheet
    #(Using JTL as make shift worksheet_name)
    def __enter_speeds_to_file__(self,file,download,upload,location,worksheet_name):

        #setting self.time to evening
        #remove on production stage
        self.setTimeEvening()
        
        #fake worksheet_name used (JTL)
        workbook=xlsxwriter.Workbook(file)
        worksheet=workbook.get_worksheet_by_name('JTL')

        #USING FAKE LOCATION DEFAULT OF KENYA

        #get current date
        date=str(datetime.datetime.now().strftime("%m/%d/%Y"))

        #Store collumn letters
        date_collumns=['B','O']#[DAY,NIGHT]
        kenya_morning=['I','J']
        kenya_evening=['V','W']

        if(self.time=='evening'):
            cell_no=int(datetime.datetime.now().strftime("%d"))+5
            #worksheet.write(str(date_collumns[1]+'%d')%(cell_no),date)
            

    #get speeds in Kenya
    def getSpeedsByInKenya(self):
        self.__kenya_server_ids=self.__get_country_servers_by_id__(self.kenyan_servers)
        self.s=speedtest.Speedtest()
        self.s.get_servers(self.__kenya_server_ids)
        self.s.get_best_server()
        self.s.download()
        self.s.upload()
        print(self.s.results.share())
        print("DOWNLOAD: ",self.__bytes_to_megabytes__(self.s.results.download))
        print("UPLOAD: ",self.__bytes_to_megabytes__(self.s.results.upload))

        file=check_or_createFile()
        print('adding download and uploads to File: '+file)
        download=self.__bytes_to_megabytes__(self.s.results.download)
        upload=self.__bytes_to_megabytes__(self.s.results.upload)

        #fill up sheet(USING JTL AS TEST)
        #self.__enter_speeds_to_file__(file,download,upload,'kenya','JTL')

    #get speeds in UK
    def getSpeedsByInUK(self):
        self.__uk_server_ids=self.__get_country_servers_by_id__(self.uk_servers)
        self.s=speedtest.Speedtest()
        self.s.get_servers(self.__uk_server_ids)
        self.s.get_best_server()
        self.s.download()
        self.s.upload()
        print(self.s.results.share())
        print("DOWNLOAD: ",self.__bytes_to_megabytes__(self.s.results.download))
        print("UPLOAD: ",self.__bytes_to_megabytes__(self.s.results.upload))

    #get speeds in USA
    def getSpeedsByInUS(self):
        self.__usa_server_ids=self.__get_country_servers_by_id__(self.usa_servers)
        self.s=speedtest.Speedtest()
        self.s.get_servers(self.__usa_server_ids)
        self.s.get_best_server()
        self.s.download()
        self.s.upload()
        print(self.s.results.share())
        print("DOWNLOAD: ",self.__bytes_to_megabytes__(self.s.results.download))
        print("UPLOAD: ",self.__bytes_to_megabytes__(self.s.results.upload))

    #get speeds in Russia
    def getSpeedsByInRussia(self):
        self.__russia_server_ids=self.__get_country_servers_by_id__(self.russia_servers)
        self.s=speedtest.Speedtest()
        self.s.get_servers(self.__russia_server_ids)
        self.s.get_best_server()
        self.s.download()
        self.s.upload()
        print(self.s.results.share())
        print("DOWNLOAD: ",self.__bytes_to_megabytes__(self.s.results.download))
        print("UPLOAD: ",self.__bytes_to_megabytes__(self.s.results.upload))
