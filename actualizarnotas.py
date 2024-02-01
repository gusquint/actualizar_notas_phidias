import csv,gspread,time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from decouple import config


class Group_Students:
    def __init__(self,name_drive,worksheet,cell_range,link_phidias,name,lugar,driver):
        if lugar=="colegio":
            sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
            PATH="C:/Users/gquintero/Desktop/python/chromedriver.exe"    
        else:
            sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
            PATH="C:/Users/PC/Desktop/python/chromedriver.exe" 
                 
        self.driver=driver
        self.name=name    
        self.sa=sa
        self.name_drive=name_drive
        self.worksheet=worksheet
        self.cell_range=cell_range
        self.link_phidias=link_phidias
        self.lugar=lugar

        #get cell range 
        google_sheet=sa.open(self.name_drive)
        hoja=google_sheet.worksheet(self.worksheet)
        cell_range_begin,cell_range_end=self.cell_range.split(":")
        cell_range_begin_letter=cell_range_begin[0]
        cell_range_begin_num=int(cell_range_begin[1:])
        cell_range_end_letter=cell_range_end[0]
        cell_range_end_num=int(cell_range_end[1:])

        #get titulos and grades
        self.headers=hoja.get(f"B3:{cell_range_end_letter}3")[0]
        self.students_info=hoja.get(cell_range)

        #get list of criteria names
        criterios=hoja.get(f"D2:{cell_range_end_letter}2")[0]
        criteria_list=[]
        for criterio in criterios:
            if criterio!="":
                criteria_list.append(criterio)        
        self.criteria=criteria_list

        #get number of grades in each criteria
        range_cells=len(self.headers)-2
        while len(criterios)<range_cells:
            criterios.append("")
        counter=1
        grades_per_criteria=[]
        for criterio in criterios[1:]:
            if criterio=="":
                counter+=1
            else:
                grades_per_criteria.append(counter)
                counter=1
        grades_per_criteria.append(counter)
        self.grades_per_criteria=grades_per_criteria


    def write_files(self):
        self.activities_to_create=[]#needed for the creation of activities that have not grades yet
        self.files_list=[]

        #append titles 
        for i in range(len(self.grades_per_criteria)):
            file_info_rows=[]
            auxiliar_list=[]
            for info in self.headers[:2]:   
                auxiliar_list.append(info)            
            begin=2+sum(self.grades_per_criteria[0:i])
            end=begin+self.grades_per_criteria[i]
            for activity in self.headers[begin:end]:
                auxiliar_list.append(activity)
            file_info_rows.append(auxiliar_list)

            grades_per_activity=[0 for _ in range(begin,end)] #needed for the creation of activities that have not grades yet
            #append names and grades
            for student in self.students_info:  
                auxiliar_list=[]  
                for value in student[:2]:#ref and name
                    auxiliar_list.append(value)
                for grade in student[begin:end]:#grades from begin to end
                    auxiliar_list.append(grade)
                file_info_rows.append(auxiliar_list)

                #append the criteria and activity name for those that have no grades yet
                if len(student[begin:end])!=0:
                    for j,grade in enumerate(student[begin:end]):
                        if grade!="":
                            grades_per_activity[j]+=1
            for j in range(len(grades_per_activity)):
                if grades_per_activity[j]==0:
                    self.activities_to_create.append((i,self.headers[begin:end][j]))#i: criteria, self.headers[begin:end][j]: title for position j

            #list of names of files
            if self.lugar=="colegio":
                file=f"C:/Users/gquintero/Desktop/python/plantillas/{self.name} {self.criteria[i]}.csv"
            else:
                file=f"C:/Users/PC/Desktop/python/plantillas/{self.name} {self.criteria[i]}.csv"
            self.files_list.append(file)

            #create file
            with open (file,"w",newline="") as file:
                filewriter=csv.writer(file,delimiter=";")
                for row in file_info_rows:
                    filewriter.writerow(row)
    
    def update(self):        
        #find the links and appends them to list        
        self.driver.get(self.link_phidias)
        links_driver=self.driver.find_elements(By.XPATH, value="//a[@class='small academic evaluation delete']")
        list_links=[]
        for link in links_driver: 
            list_links.append(link.get_attribute("href"))

        #deletes all the values
        for link in list_links:
            self.driver.get(link)

        #get the upload links
        temporary_list=self.driver.find_elements(By.XPATH, value="//a[@class='small xls']")
        link_upload_list=[]
        for element in temporary_list:
            link_upload_list.append(element.get_attribute("href"))

        #upload files
        for i in range(len(link_upload_list)):
            self.driver.get(link_upload_list[i])
            upload_file = self.driver.find_element(By.XPATH, value="//input[@type='file']")
            upload_file.send_keys(self.files_list[i])
            self.driver.find_element(By.XPATH, value="//input[@type='submit']").click()
        
        #create activities that have no grades
        add_evaluation_link=self.driver.find_elements(By.XPATH, value="//a[@class='small add academic evaluation']")
        links=[element.get_attribute("href") for element in add_evaluation_link]

        for criteria,activity in self.activities_to_create:
            self.driver.get(links[criteria])
            search=self.driver.find_element(By.NAME,value="name")
            search.send_keys(activity)
            search.send_keys(Keys.RETURN)


        


def open_phidias(link,lugar,ussername,password):
    if lugar=="colegio":
        PATH="C:/Users/gquintero/Desktop/python/chromedriver.exe"
    else:
        PATH="C:/Users/PC/Desktop/python/chromedriver.exe" 
    driver=webdriver.Chrome(PATH)
    driver.get(link)
    search=driver.find_element(By.ID,value="autofocus")
    search.send_keys(ussername)
    search=driver.find_element(By.NAME,value="_height")
    search.send_keys(password)
    search.send_keys(Keys.RETURN)
    return driver


def close_phidias(driver,seconds):
    time.sleep(seconds)
    driver.close()


def subir_notas_phidias(name,drive_name,worksheet,range,lugar,driver,link):
    group=Group_Students(drive_name,worksheet,range,link,name,lugar,driver)
    group.write_files()
    group.update()
    print(f"Las notas de {name} fueron acualizadas en phidias exitosamente")




def main():
    lugar="colegio"
    driver=open_phidias("https://colegionuevayork.phidias.co/person/employee/dashboard/classes?person=3211",lugar,config("phidias_usuario"),config("phidias_password"))    

    subir_notas_phidias("9B","personal notas 9 tercer periodo","Grades","B4:K26",lugar,driver,"https://colegionuevayork.phidias.co/academic/grading/period?group=3391&period=15&person=3211&_ach=fef2c0373e53f73640cd6354f07df32e")
    subir_notas_phidias("9C","personal notas 9 tercer periodo","Grades","B28:K50",lugar,driver,"https://colegionuevayork.phidias.co/academic/grading/period?group=3392&period=15&_ach=608b295d5fca505659e931e38250435c")
    subir_notas_phidias("10MAI","Personal notas 10 MAI tercer periodo","Grades","B4:N25",lugar,driver,"https://colegionuevayork.phidias.co/academic/grading/period?group=3826&period=15&person=3211&_ach=f071342ce6622b408ddf608147ca1e71")
    subir_notas_phidias("11MAI","Personal notas 11 MAI tercer periodo","Grades","B4:K23",lugar,driver,"https://colegionuevayork.phidias.co/academic/grading/period?group=3830&period=15&person=3211&_ach=1edfde19181b9715860d36d3653adfbe")

    close_phidias(driver,0)



if __name__ == "__main__":
    main()