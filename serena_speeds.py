import speedtest

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


    #get speeds in Kenya
    def getSpeedsByInKenya(self):
        self.__kenya_server_ids=self.__get_country_servers_by_id__(self.kenyan_servers)
        self.s.get_servers(self.__kenya_server_ids)
        self.s.get_best_server()
        self.s.download()
        self.s.upload()
        print(self.s.results.share())
        print("DOWNLOAD: ",self.__bytes_to_megabytes__(self.s.results.download))
        print("UPLOAD: ",self.__bytes_to_megabytes__(self.s.results.upload))


