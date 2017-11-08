# -*- coding: utf-8 -*-
import xlwt
import os



# kodowanie arkusza
#book = xlwt.Workbook(encoding="utf-8")

# tworzymy dowolną ilość arkuszy (zakładek)
#sheet1 = book.add_sheet("punkty")


# umieszczamy w nich dane
print ("przed wpisaniem nazwy pliku wklej go")
print ("do tego samego katalogu co ten plik wykonawczy")
#nazwa_pliku_wejscia = input("Podaj nazwe pliku: ")
#plik = open(nazwa_pliku_wejscia)
#plik1 = open("robot\s12proc02.ls")
#plik = open("robot\s12proc01.ls")
katalog= (os.listdir('robot'))
print(katalog)
#nr_pliku = katalog.index('s12proc05.ls')

nazwa1_programu_1= "s12proc0"
nazwa1_programu_3= ".ls"

for j in range(1,9):
        
        nazwa1_programu=nazwa1_programu_1 + str(j) + nazwa1_programu_3
        
        print(nazwa1_programu in katalog)
        if nazwa1_programu in katalog:

                nr_pliku = katalog.index(nazwa1_programu)
        
                plik = open("robot" + "\\" + katalog[nr_pliku] )
                print(plik)
           #     try:
       
                tekst = plik.read()

               # finally:
             #           plik.close()

                #print tekst
                book = xlwt.Workbook(encoding="utf-8")
                sheet1 = book.add_sheet("punkty")
                sheet1.write(0, 0, "Punkt")
                sheet1.write(0, 1, "Styl")
                sheet1.write(0, 3, "Nazwa")
                sheet1.write(0, 4, "Robot ID")
                robot=tekst[tekst.find("ROBOT"):tekst.find("ROBOT")+50]
                sheet1.write(1, 4, robot[robot.find(":")+1:robot.find(";")-1])
                print ("--------------------------")
                fraza1 ='P['
                fraza3=":"
                i=0
                for x in range(0,len(tekst)):
                    fraza= fraza1+ str(x)+ fraza3
                    y= tekst.find(fraza)
                    if y!=-1 :
                       print (fraza)
                       znalazl= tekst[tekst.find(fraza):tekst.find(']', tekst.find(fraza), tekst.find(fraza)+30)]
                       znalaz= znalazl[5:]
                       print (znalaz)
                       i+=1
                       punkt=znalazl[:znalazl.find(':')]
                       nazwa_punktu=znalazl[znalazl.find(':')+1:]
                       sheet1.write(i, 0, punkt+"]")
                       sheet1.write(i, 1, nazwa_punktu)
                       sheet1.write(i, 3, nazwa_punktu[:6]) 
                #nazwa_pliku_wyjscia = nazwa_pliku_wejscia[:-3] + ".xls"
                # zapisujemy do pliku [:punkt.find(']')]
                       #book.save(nazwa_pliku_wyjscia)
                book.save('excel\\'+ nazwa1_programu[:-3] + ".xls")
                plik.close()

nazwa1_programu_1= "s12proc"
nazwa1_programu_3= ".ls"

for j in range(1,9):
        
        nazwa1_programu=nazwa1_programu_1 + str(j) + nazwa1_programu_3
        
        print(nazwa1_programu in katalog)
        if nazwa1_programu in katalog:

                nr_pliku = katalog.index(nazwa1_programu)
        
                plik = open("robot" + "\\" + katalog[nr_pliku] )
                print(plik)
           #     try:
       
                tekst = plik.read()

               # finally:
             #           plik.close()

                #print tekst
                book = xlwt.Workbook(encoding="utf-8")
                sheet1 = book.add_sheet("punkty")
                sheet1.write(0, 0, "Punkt")
                sheet1.write(0, 1, "Styl")
                sheet1.write(0, 3, "Nazwa")
                sheet1.write(0, 4, "Robot ID")
                robot=tekst[tekst.find("ROBOT"):tekst.find("ROBOT")+50]
                sheet1.write(1, 4, robot[robot.find(":")+1:robot.find(";")-1])
                print ("--------------------------")
                fraza1 ='P['
                fraza3=":"
                i=0
                for x in range(0,len(tekst)):
                    fraza= fraza1+ str(x)+ fraza3
                    y= tekst.find(fraza)
                    if y!=-1 :
                       print (fraza)
                       znalazl= tekst[tekst.find(fraza):tekst.find(']', tekst.find(fraza), tekst.find(fraza)+30)]
                       znalaz= znalazl[5:]
                       print (znalaz)
                       i+=1
                       punkt=znalazl[:znalazl.find(':')]
                       nazwa_punktu=znalazl[znalazl.find(':')+1:]
                       sheet1.write(i, 0, punkt+"]")
                       sheet1.write(i, 1, nazwa_punktu)
                       sheet1.write(i, 3, nazwa_punktu[:6]) 
                #nazwa_pliku_wyjscia = nazwa_pliku_wejscia[:-3] + ".xls"
                # zapisujemy do pliku [:punkt.find(']')]
                       #book.save(nazwa_pliku_wyjscia)
                book.save('excel\\'+ nazwa1_programu[:-3] + ".xls")
                plik.close()
