import openpyxl

import sys
from PyQt5.QtWidgets import (QWidget, QTableWidget, QHBoxLayout, QApplication, QTableWidgetItem )
from PyQt5.QtGui     import QBrush, QColor #,  QFont 

def cloths():   
    book = openpyxl.load_workbook('athleisure.xlsx')
    sheet = book.active

    count = 2
    index = 0
    
    brand_list = []
    style_list = []
    price_list = []
    tech_list = []
    rating_list = []
    reigon_list = []
    sex_list = []
    

    brand = sheet['A2']
    style = sheet['B2']
    price = sheet['C2']
    tech = sheet['D2']
    rating = sheet['E2']
    reigon = sheet['F2']
    sex = sheet['G2']

    brand_list += [brand.value]
    style_list += [style.value]
    price_list += [float(price.value)]
    tech_list += [tech.value]
    rating_list += [float(str(rating.value)[-3:])]
    reigon_list += [reigon.value]
    sex_list += [sex.value]

    while brand.value != None:
        count += 1
        brand = sheet['A' + str(count)]
        style = sheet['B' + str(count)]
        price = sheet['C' + str(count)]
        tech = sheet['D' + str(count)]
        rating = sheet['E' + str(count)]
        reigon = sheet['F' + str(count)]
        sex = sheet['G' + str(count)]
        
        brand_list += [brand.value]
        style_list += [style.value]
        if price.value == None:
            price_list += [float(0)]
        else:
            price_list += [float(price.value)]           
        tech_list += [tech.value]
        if rating.value == None:
            rating_list += [0]
        else:
            rating_list += [float(str(rating.value)[-3:])]
        reigon_list += [reigon.value]
        sex_list += [sex.value]

    return [brand_list,style_list,price_list,tech_list,rating_list,reigon_list,sex_list]


def weight():
    price = [i * 0.6 for i in cloths()[2]]
    rating = [j * 0.4 * 10 for j in cloths()[4]]

    return [x + y for x, y in zip(price, rating)]

    
def bubbleSortW(arr,w):
    n = len(arr[2])
 
    # Traverse through all array elements
    for i in range(n):
 
        # Last i elements are already in place
        for j in range(0, n-i-1):
 
            # traverse the array from 0 to n-i-1
            # Swap if the element found is greater
            # than the next element
            if w[j] < w[j+1] :
                w[j], w[j+1] = w[j+1], w[j]
                arr[2][j], arr[2][j+1] = arr[2][j+1], arr[2][j]
                arr[0][j], arr[0][j+1] = arr[0][j+1], arr[0][j]
                arr[1][j], arr[1][j+1] = arr[1][j+1], arr[1][j]
                arr[3][j], arr[3][j+1] = arr[3][j+1], arr[3][j]
                arr[4][j], arr[4][j+1] = arr[4][j+1], arr[4][j]
                arr[5][j], arr[5][j+1] = arr[5][j+1], arr[5][j]
                arr[6][j], arr[6][j+1] = arr[6][j+1], arr[6][j]
    return arr

    

class Table(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Top Brands by Reigon")
        self.resize(660,300 );
        conLayout = QHBoxLayout()

        tableWidget = QTableWidget()
        tableWidget.setRowCount(21)
        tableWidget.setColumnCount(16)
        conLayout.addWidget(tableWidget)

        # Hide headers
        tableWidget.horizontalHeader().setVisible(False)
        tableWidget.verticalHeader().setVisible(False)

        #tableWidget.setHorizontalHeaderLabels(['Column1','Column1','Column1'])  

        # Sets the span of the table element at (row , column ) to the number of rows 
        # and columns specified by (rowSpanCount , columnSpanCount ).
        
#-------------------------------------------------------#
        tableWidget.setSpan(0, 1, 1, 16) 
        newItem = QTableWidgetItem("North America")  
        tableWidget.setItem(0, 1, newItem) 

        tableWidget.setSpan(1, 1, 1, 5)   
        newItem = QTableWidgetItem("Best")  
        tableWidget.setItem(1, 1, newItem)  

        tableWidget.setSpan(1, 6, 1, 5)   
        newItem = QTableWidgetItem("Better")  
        tableWidget.setItem(1, 6, newItem)

        tableWidget.setSpan(1, 11, 1, 5)   
        newItem = QTableWidgetItem("Good")  
        tableWidget.setItem(1, 11, newItem)
#-------------------------------------------------------#
        tableWidget.setSpan(7, 1, 1, 16) 
        newItem = QTableWidgetItem("Europe")  
        tableWidget.setItem(7, 1, newItem)
        
        tableWidget.setSpan(8, 1, 1, 5)   
        newItem = QTableWidgetItem("Best")  
        tableWidget.setItem(8, 1, newItem)  

        tableWidget.setSpan(8, 6, 1, 5)   
        newItem = QTableWidgetItem("Better")  
        tableWidget.setItem(8, 6, newItem)

        tableWidget.setSpan(8, 11, 1, 5)   
        newItem = QTableWidgetItem("Good")  
        tableWidget.setItem(8, 11, newItem)
#-------------------------------------------------------#
        tableWidget.setSpan(14, 1, 1, 16) 
        newItem = QTableWidgetItem("Asia")  
        tableWidget.setItem(14, 1, newItem)
        
        tableWidget.setSpan(15, 1, 1, 5)   
        newItem = QTableWidgetItem("Best")  
        tableWidget.setItem(15, 1, newItem)  

        tableWidget.setSpan(15, 6, 1, 5)   
        newItem = QTableWidgetItem("Better")  
        tableWidget.setItem(15, 6, newItem)

        tableWidget.setSpan(15, 11, 1, 5)   
        newItem = QTableWidgetItem("Good")  
        tableWidget.setItem(15, 11, newItem)

#-------------------------------------------------------#

        newItem = QTableWidgetItem("Reigon")  
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(0, 0, newItem)
        

        newItem = QTableWidgetItem("Type") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(1, 0, newItem)

        newItem = QTableWidgetItem("Top Brands")  
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(2, 0, newItem)  

        newItem = QTableWidgetItem("Top Styles") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(3, 0, newItem)

        newItem = QTableWidgetItem("Tech")  
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(4, 0, newItem)  

        newItem = QTableWidgetItem("Rating") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(5, 0, newItem)

        newItem = QTableWidgetItem("Price") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(6, 0, newItem)

        newItem = QTableWidgetItem("Reigon")  
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(7, 0, newItem)
        

        newItem = QTableWidgetItem("Type") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(8, 0, newItem)

        newItem = QTableWidgetItem("Top Brands")  
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(9, 0, newItem)  

        newItem = QTableWidgetItem("Top Styles") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(10, 0, newItem)

        newItem = QTableWidgetItem("Tech")  
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(11, 0, newItem)  

        newItem = QTableWidgetItem("Rating") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(12, 0, newItem)

        newItem = QTableWidgetItem("Price") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(13, 0, newItem)

        newItem = QTableWidgetItem("Reigon")  
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(14, 0, newItem)
        

        newItem = QTableWidgetItem("Type") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(15, 0, newItem)

        newItem = QTableWidgetItem("Top Brands")  
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(16, 0, newItem)  

        newItem = QTableWidgetItem("Top Styles") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(17, 0, newItem)

        newItem = QTableWidgetItem("Tech")  
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(18, 0, newItem)  

        newItem = QTableWidgetItem("Rating") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(19, 0, newItem)

        newItem = QTableWidgetItem("Price") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(20, 0, newItem)
        

        [best1,best2,best3] = [0,0,0]
        [better1,better2,better3] = [0,0,0]
        [good1,good2,good3] = [0,0,0]

        reigon = bubbleSortW(cloths(),weight())[5]
        brand = bubbleSortW(cloths(),weight())[0]
        style = bubbleSortW(cloths(),weight())[1]
        price = bubbleSortW(cloths(),weight())[2]
        tech = bubbleSortW(cloths(),weight())[3]
        rating = bubbleSortW(cloths(),weight())[4]

        #[brand_list,style_list,price_list,tech_list,rating_list,reigon_list,sex_list]
        
        for i in range(len(reigon)):
            if reigon[i] == 'NA':
                item0 = QTableWidgetItem(brand[i])
                item1 = QTableWidgetItem(style[i])
                item2 = QTableWidgetItem(tech[i])
                item3 = QTableWidgetItem(str(rating[i]))
                item4 = QTableWidgetItem('$ '+str(price[i]))
                if price[i] > 120 and best1 != 5:
                    tableWidget.setItem(2, 1 + best1, item0)
                    tableWidget.setItem(3, 1 + best1, item1)
                    tableWidget.setItem(4, 1 + best1, item2)
                    tableWidget.setItem(5, 1 + best1, item3)
                    tableWidget.setItem(6, 1 + best1, item4)
                    best1 += 1
                elif (price[i] > 65 and price[i] < 120) and better1 != 5:
                    tableWidget.setItem(2, 6 + better1, item0)
                    tableWidget.setItem(3, 6 + better1, item1)
                    tableWidget.setItem(4, 6 + better1, item2)
                    tableWidget.setItem(5, 6 + better1, item3)
                    tableWidget.setItem(6, 6 + better1, item4)
                    better1 += 1
                elif price[i] < 65 and good1 != 5:
                    tableWidget.setItem(2, 11 + good1, item0)
                    tableWidget.setItem(3, 11 + good1, item1)
                    tableWidget.setItem(4, 11 + good1, item2)
                    tableWidget.setItem(5, 11 + good1, item3)
                    tableWidget.setItem(6, 11 + good1, item4)
                    good1 += 1

            if reigon[i] == 'EU':
                item0 = QTableWidgetItem(brand[i])
                item1 = QTableWidgetItem(style[i])
                item2 = QTableWidgetItem(tech[i])
                item3 = QTableWidgetItem(str(rating[i]))
                item4 = QTableWidgetItem('$ '+str(price[i]))
                if price[i] > 120 and best2 != 5:
                    tableWidget.setItem(9, 1 + best2, item0)
                    tableWidget.setItem(10, 1 + best2, item1)
                    tableWidget.setItem(11, 1 + best2, item2)
                    tableWidget.setItem(12, 1 + best2, item3)
                    tableWidget.setItem(13, 1 + best2, item4)
                    best2 += 1
                elif (price[i] > 65 and price[i] < 120) and better2 != 5:
                    tableWidget.setItem(9, 6 + better2, item0)
                    tableWidget.setItem(10, 6 + better2, item1)
                    tableWidget.setItem(11, 6 + better2, item2)
                    tableWidget.setItem(12, 6 + better2, item3)
                    tableWidget.setItem(13, 6 + better2, item4)
                    better2 += 1
                elif price[i] < 65 and good2 != 5:
                    tableWidget.setItem(9, 11 + good2, item0)
                    tableWidget.setItem(10, 11 + good2, item1)
                    tableWidget.setItem(11, 11 + good2, item2)
                    tableWidget.setItem(12, 11 + good2, item3)
                    tableWidget.setItem(13, 11 + good2, item4)
                    good2 += 1
                    
            if reigon[i] == 'AS':
                item0 = QTableWidgetItem(brand[i])
                item1 = QTableWidgetItem(style[i])
                item2 = QTableWidgetItem(tech[i])
                item3 = QTableWidgetItem(str(rating[i]))
                item4 = QTableWidgetItem('$ '+str(price[i]))
                if price[i] > 120 and best3 != 5:
                    tableWidget.setItem(16, 1 + best3, item0)
                    tableWidget.setItem(17, 1 + best3, item1)
                    tableWidget.setItem(18, 1 + best3, item2)
                    tableWidget.setItem(19, 1 + best3, item3)
                    tableWidget.setItem(20, 1 + best3, item4)
                    best3 += 1
                elif (price[i] > 65 and price[i] < 120) and better3 != 5:
                    tableWidget.setItem(16, 6 + better3, item0)
                    tableWidget.setItem(17, 6 + better3, item1)
                    tableWidget.setItem(18, 6 + better3, item2)
                    tableWidget.setItem(19, 6 + better3, item3)
                    tableWidget.setItem(20, 6 + better3, item4)
                    better3 += 1
                elif price[i] < 65 and good3 != 5:
                    tableWidget.setItem(16, 11 + good3, item0)
                    tableWidget.setItem(17, 11 + good3, item1)
                    tableWidget.setItem(18, 11 + good3, item2)
                    tableWidget.setItem(19, 11 + good3, item3)
                    tableWidget.setItem(20, 11 + good3, item4)
                    good3 += 1


  
        self.setLayout(conLayout)

class Top(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Top Ten Styles")
        self.resize(660,300 );
        conLayout = QHBoxLayout()

        tableWidget = QTableWidget()
        tableWidget.setRowCount(12)
        tableWidget.setColumnCount(6)
        conLayout.addWidget(tableWidget)

        # Hide headers
        tableWidget.horizontalHeader().setVisible(False)
        tableWidget.verticalHeader().setVisible(False)

        #tableWidget.setHorizontalHeaderLabels(['Column1','Column1','Column1'])  

        # Sets the span of the table element at (row , column ) to the number of rows 
        # and columns specified by (rowSpanCount , columnSpanCount ).
        
#-------------------------------------------------------#
        inp = input("Enter brand to search: ")
        
        tableWidget.setSpan(0, 0, 1, 6) 
        newItem = QTableWidgetItem(inp)  
        tableWidget.setItem(0, 0, newItem)

        for i in range(10):
            newItem = QTableWidgetItem(str(i+1))  
            tableWidget.setItem(i+2, 0, newItem)

        newItem = QTableWidgetItem("Rank")
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(1, 0, newItem)

        newItem = QTableWidgetItem("Type")
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(1, 1, newItem)

        newItem = QTableWidgetItem("Price")
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(1, 2, newItem)

        newItem = QTableWidgetItem("Rating")
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(1, 3, newItem)

        newItem = QTableWidgetItem("Tech")
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(1, 4, newItem)

        newItem = QTableWidgetItem("Reigon")
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(1, 5, newItem)

        reigon = bubbleSortW(cloths(),weight())[5]
        brand = bubbleSortW(cloths(),weight())[0]
        style = bubbleSortW(cloths(),weight())[1]
        price = bubbleSortW(cloths(),weight())[2]
        tech = bubbleSortW(cloths(),weight())[3]
        rating = bubbleSortW(cloths(),weight())[4]

        #[brand_list,style_list,price_list,tech_list,rating_list,reigon_list,sex_list]
        count = 0
        
        for i in range(len(reigon)):
            if brand[i] == inp or brand[i] == inp + ' ':
                item0 = QTableWidgetItem(style[i])
                item1 = QTableWidgetItem('$ '+str(price[i]))
                item2 = QTableWidgetItem(str(rating[i]))
                item3 = QTableWidgetItem(tech[i])
                if reigon[i] == 'NA':
                    r = "North America"
                elif reigon[i] == 'EU':
                    r = "Europe"
                else:
                    r = "Asia"
                item4 = QTableWidgetItem(r)
                tableWidget.setItem(2+count,1 , item0)
                tableWidget.setItem(2+count,2 , item1)
                tableWidget.setItem(2+count,3 , item2)
                tableWidget.setItem(2+count,4 , item3)
                tableWidget.setItem(2+count,5 , item4)
                count += 1
#-------------------------------------------------------#
        '''
        newItem = QTableWidgetItem("Rating") 
        newItem.setForeground(QBrush(QColor(255, 0, 0)))        
        tableWidget.setItem(12, 0, newItem)


        

        count1 = 0
        count2 = 0
        count3 = 0

        reigon = bubbleSort(cloths())[5]
        brand = bubbleSort(cloths())[0]
        style = bubbleSort(cloths())[1]
        price = bubbleSort(cloths())[2]
        tech = bubbleSort(cloths())[3]
        rating = bubbleSort(cloths())[4]

        #[brand_list,style_list,price_list,tech_list,rating_list,reigon_list,sex_list]
        '''

        self.setLayout(conLayout)
        
def displayTable():
    app = QApplication(sys.argv)
    example = Table()  
    example.show()   
    sys.exit(app.exec_())

def search():
    app = QApplication(sys.argv)
    example = Top()  
    example.show()   
    sys.exit(app.exec_())

       
def main():
        print("---------------------------------------------------------")
        print("WELCOME TO AMART (Apparel Market Analysis Research Tool)")
        print("---------------------------------------------------------")
        print("Follow the instructions below to use ")
        print()
        print("- To show table of brands sorted by reigon, type 'displayTable()' in the next line")
        print("- To search brands, type 'search()' in the next line")
        print("- To restart type 'main()'")

print("Type 'main()' to start")

            

    
    
                

    


        
        
        

