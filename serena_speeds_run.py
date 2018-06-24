import serena_speeds

#hold isp
isp=0

def prompt_day_or_night():
    print('Please select time of speed test')
    print('1 - morning')
    print('2 - evening')
    print('Other - Exit application')
    time=int(input('Enter Here: '))
    
    if(time==1):
        prompt_select_isp_in_use(1)
    elif(time==2):
        prompt_select_isp_in_use(2)
    else:
        print('Sorry, restart application and choose a valid option')
        input()
        exit()

def prompt_select_isp_in_use(time):
    print('Please choose the ISP currenctly connect to from List:')
    print('1 - JTL (Jamii Telkom)')
    print('2 - SAF (Safaricom)')
    print('3 - LQD (Liquid Telkom)')
    print('4 - Exit Application')
    isp=int(input('Enter Here: '))

    if(isp==4):
        print('Exiting Application, Press[ENTER] to continue')
        input()
        exit()
    elif(isp==1):
        print('Starting JTL Test')
        serena=serena_speeds.SerenaSpeedTester()
        serena.setWorksheetJTL()

        #set time
        if(time==1):
            serena.setTimeMorning()
            print('Morning, Set')
        elif(time==2):
            serena.setTimeEvening()
            print('Evening, Set')
            
        serena.getSpeedsByInKenya()
        serena.getSpeedsByInRussia()
        serena.getSpeedsByInUK()
        serena.getSpeedsByInUS()
        print('Done with JTL')
        prompt_select_isp_in_use(time)
    elif(isp==2):
        print('Starting SAF Test')
        serena=serena_speeds.SerenaSpeedTester()
        serena.setWorkSheetSAF()
        
        #set time
        if(time==1):
            serena.setTimeMorning()
        elif(time==2):
            serena.setTimeEvening()
        
        serena.getSpeedsByInKenya()
        serena.getSpeedsByInRussia()
        serena.getSpeedsByInUK()
        serena.getSpeedsByInUS()
        print('Done with SAF')
        prompt_select_isp_in_use(time)
    elif(isp==3):
        print('Starting LQD Test')
        serena=serena_speeds.SerenaSpeedTester()
        serena.setWorkSheetLQD()

        #set time
        if(time==1):
            serena.setTimeMorning()
        elif(time==2):
            serena.setTimeEvening()

        serena.getSpeedsByInKenya()
        serena.getSpeedsByInRussia()
        serena.getSpeedsByInUK()
        serena.getSpeedsByInUS()
        print('Done with LQD')
        prompt_select_isp_in_use(time)
    else:
        print('please select a valid option, or quit')
        prompt_select_isp_in_use(time)



def main():
    prompt_day_or_night()

if __name__=='__main__':
    main()
